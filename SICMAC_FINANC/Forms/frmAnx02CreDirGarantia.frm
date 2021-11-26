VERSION 5.00
Begin VB.Form frmAnx02CreDirGarantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexo 2: Créditos Directos e Indirectos por Tipo de Garantía"
   ClientHeight    =   930
   ClientLeft      =   2295
   ClientTop       =   2340
   ClientWidth     =   4470
   Icon            =   "frmAnx02CreDirGarantia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAnx02CreDirGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nCuentas As Integer
Dim dFecha As Date
Dim sSql As String
Dim lnTpoCambio As Currency
Dim lsTitulo As String

Dim R As New ADODB.Recordset
Dim oCon As New DConecta 'NAGL 20171211 Agregó New, para instanciar

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oBarra As clsProgressBar
Dim nTFilaActual As Integer
Dim nTColActual As Integer

Dim sservidorconsolidada As String
Dim cVigente As String
Dim cPigno As String
Private Sub TotalesExcel(nFila As Integer)
'16 17 19 21 22 23
xlHoja1.Range("AD" & 17 + nFila & ":AD" & 17 + nFila).Formula = "=SUM(D" & 17 + nFila & ":AC" & 17 + nFila & ")"
xlHoja1.Range("AD" & 18 + nFila & ":AD" & 18 + nFila).Formula = "=SUM(D" & 18 + nFila & ":AC" & 18 + nFila & ")"
xlHoja1.Range("AD" & 20 + nFila & ":AD" & 20 + nFila).Formula = "=SUM(D" & 20 + nFila & ":AC" & 20 + nFila & ")"
xlHoja1.Range("AD" & 22 + nFila & ":AD" & 22 + nFila).Formula = "=SUM(D" & 22 + nFila & ":AC" & 22 + nFila & ")"
xlHoja1.Range("AD" & 23 + nFila & ":AD" & 23 + nFila).Formula = "=SUM(D" & 23 + nFila & ":AC" & 23 + nFila & ")"
xlHoja1.Range("AD" & 24 + nFila & ":AD" & 24 + nFila).Formula = "=SUM(D" & 24 + nFila & ":AC" & 24 + nFila & ")"

'Total Vigentes
Dim nCol As Integer
Dim sCol As String
For nCol = 4 To 30
    sCol = ExcelColumnaString(nCol)
  '  If xlHoja1.Cells(17 + nFila, nCol) <> "" Or xlHoja1.Cells(18 + nFila, nCol) <> "" Then
   '     xlHoja1.Range(sCol & 16 + nFila & ":" & sCol & 16 + nFila).Formula = "=+" & sCol & 17 + nFila & "+" & sCol & 18 + nFila
    'End If
    
    'If xlHoja1.Cells(22 + nFila, nCol) <> "" Or xlHoja1.Cells(23 + nFila, nCol) <> "" Then
     '   xlHoja1.Range(sCol & 21 + nFila & ":" & sCol & 21 + nFila).Formula = "=+" & sCol & 22 + nFila & "+" & sCol & 23 + nFila
   ' End If
        
    'If xlHoja1.Cells(16 + nFila, nCol) <> "" Or xlHoja1.Cells(19 + nFila, nCol) <> "" Or xlHoja1.Cells(20 + nFila, nCol) <> "" Or xlHoja1.Cells(21 + nFila, nCol) <> "" Or xlHoja1.Cells(24 + nFila, nCol) <> "" Then
        xlHoja1.Range(sCol & 25 + nFila & ":" & sCol & 25 + nFila).Formula = "=+" & sCol & 17 + nFila & "+" & sCol & 18 + nFila & "+" & sCol & 20 + nFila & "+" & sCol & 22 + nFila & "+" & sCol & 23 + nFila & "+" & sCol & 24 + nFila
   ' End If
Next
End Sub

Private Sub GeneraReporteAnexo2(pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String)
Dim I As Integer
Dim k As Integer
Dim j As Integer
Dim nFila As Integer
Dim nIni  As Integer
Dim lNegativo As Boolean
Dim ldFecha   As Date
Dim sConec As String
Dim sTipoGara As String
Dim nTipoCred As Integer
Dim lsArchivo As String
Dim lsArchivo1 As String
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean
Dim RutaArchivo As String
Dim fs As Scripting.FileSystemObject
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
Dim rs As ADODB.Recordset

    ldFecha = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1

    lsArchivo = "Anexo02Garantias"
    lsArchivo1 = "\spooler\Anx02_" & Format(gdFecha, "YYYYMM") & gsCodUser & ".xls"
    lsNomHoja = "Anexo 2"
    RutaArchivo = App.path & "\FormatoCarta\" & lsArchivo & ".xls"
    If fs.FileExists(RutaArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    If oCon.AbreConexion = False Then Exit Sub
     
    oBarra.Max = 14
    oBarra.Progress 0, lsTitulo, "Generando reporte...", , vbBlue
    
    For Each xlHoja1 In xlLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
   
    xlHoja1.Range("B5") = "AL " & Format(ldFecha, "DD") & " DE  " & UCase(Format(ldFecha, "MMMM")) & " DEL  " & Format(ldFecha, "YYYY")

    Set rs = New ADODB.Recordset
    oCon.CommadTimeOut = 7200
    
'    sSql = " SELECT SUBSTRING(CA.cTpoCredCod,1,1) TipoCred, "
'    sSql = sSql & " sum((case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & pnTpoCambio & " end)*(case cSitCtb when '1' then c.nSaldoCap when '2' then c.nSaldoCap-c.nCapVencido else 0 end)) nVigente, "
'    sSql = sSql & " sum((case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & pnTpoCambio & " end)*(case cSitCtb when '4' then c.nSaldoCap else 0 end)) nRefinanciado, "
'    sSql = sSql & " sum((case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & pnTpoCambio & " end)*(case cSitCtb when '5' then c.nSaldoCap when '2' then c.nCapVencido else 0 end)) nVencido, "
'    sSql = sSql & " sum((case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & pnTpoCambio & " end)*(case cSitCtb when '6' then c.nSaldoCap else 0 end)) nJudicial, "
'    sSql = sSql & " TipoGar = CASE WHEN CA.nGarant = 1 THEN 1 "
'    sSql = sSql & "         WHEN CA.nGarant = 3 THEN 2 "
'    sSql = sSql & "         WHEN CA.nGarant = 4 THEN 3 END "
'    sSql = sSql & " FROM DBConsolidada..CreditoConsol C "
'    sSql = sSql & "   inner join ColocCalifProv CA ON C.cCtaCod = CA.cCtaCod "
'    sSql = sSql & "   WHERE C.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107,2092) "
'    sSql = sSql & "          and C.cSitCtb in ('1','2','4','5','6') "
'    sSql = sSql & "   GROUP BY SUBSTRING(CA.cTpoCredCod,1,1),CA.nGarant "
'    sSql = sSql & " ORDER BY SUBSTRING(CA.cTpoCredCod,1,1) "
'   ALPA 20150427
'JUEZ 20160312 *********************************************************************
'  sSql = " select substring(Producto,1,1) TipoCred, "
''  sSql = sSql & " TipoGar = CASE WHEN nGarant = 1 THEN 1 "
''  sSql = sSql & "         WHEN nGarant = 3 THEN 2 "
''  sSql = sSql & "         WHEN nGarant = 4 THEN 3 END, "
'  sSql = sSql & "   TipoGar = case nTipoLeasing when 'L' then 4 else CASE WHEN nGarant = 1 THEN 1  WHEN nGarant = 3 THEN 2 WHEN nGarant = 4 THEN 3 END end,"
'  sSql = sSql & "   nVigente = sum(Case when cRefinan='N' then (case when Moneda = '1' then VigSaldo else VigSaldo*" & pnTpoCambio & " end) else 0 end), "
'  sSql = sSql & "   nRefinanciado = sum(Case when cRefinan='R' then (case when Moneda = '1' then VigSaldo else VigSaldo*" & pnTpoCambio & " end) else 0 end), "
'  sSql = sSql & "   nVencido = sum(case when Moneda = '1' then Ven1Saldo + Ven2Saldo else (Ven1Saldo + Ven2Saldo)*" & pnTpoCambio & " end), "
'  sSql = sSql & "   nJudicial = sum(case when Moneda = '1' then JudSDSaldo + JudCDSaldo else (JudSDSaldo + JudCDSaldo)*" & pnTpoCambio & " end) "
'  sSql = sSql & " from ( "
'  sSql = sSql & "   select  CP.nGarant, substring(C.cLineaCred,5,1) Moneda,"
'  sSql = sSql & "     CC.cTpoCredCod Producto, "
'  sSql = sSql & "     C.cRefinan, "
'  sSql = sSql & "     sum(C.nSaldoCap) AS TotSaldoCap, "
'  sSql = sSql & "     isnull(sum(case when C.cSitCtb IN ('1','4') then C.nSaldoCap "
'  sSql = sSql & "             when C.cSitCtb='2' then C.nSaldoCap - isnull(C.nCapVencido,0) "
'  sSql = sSql & "           end),0) VigSaldo, "
'  sSql = sSql & "     isnull(sum(case when C.cSitCtb='5' and substring(CC.cTpoCredCod,1,1) not in ('6','7','8') and C.nPrdEstado not in (2201,2205) then C.nSaldoCap "
'  sSql = sSql & "             when C.cSitCtb='2' and substring(CC.cTpoCredCod,1,1) in ('6','7','8') then isnull(C.nCapVencido,0) "
'  sSql = sSql & "           end),0) Ven1Saldo, "
'  sSql = sSql & "     isnull(sum(case when C.cSitCtb='5' and C.nPrdEstado not in (2201,2205) and substring(CC.cTpoCredCod,1,1) in ('6','7','8') then C.nSaldoCap end),0) Ven2Saldo, "
'  sSql = sSql & "     isnull(sum(case when C.cSitCtb='5' and C.nPrdEstado in (2201,2205) and isnull(C.nDemanda,2) = 2 then C.nSaldoCap  end),0) JudSDSaldo, "
'  sSql = sSql & "     isnull(sum(case when C.cSitCtb='6' and isnull(C.nDemanda,2) = 1 then C.nSaldoCap end),0) JudCDSaldo, "
'  sSql = sSql & "     nTipoLeasing=case when substring(C.cCtaCod,6,3)='515' then 'L' else 'N' end"
'  sSql = sSql & "   from " & sservidorconsolidada & "CreditoConsol C "
'  sSql = sSql & "     inner join " & sservidorconsolidada & "CreditoConsolTotal CC on (C.cCtaCod=CC.cCtaCod) "
'  sSql = sSql & "     inner join ColocCalifProv CP on (C.cCtaCod=CP.cCtaCod) "
'  sSql = sSql & "   where C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) "
'  sSql = sSql & "   group by CC.cTpoCredCod, CP.nGarant, substring(C.cLineaCred,5,1), C.cRefinan,case when substring(C.cCtaCod,6,3)='515' then 'L' else 'N' end  "
'  sSql = sSql & " ) A "
'  sSql = sSql & " group by substring(Producto,1,1), Case nTipoLeasing when 'L' then '3' else nGarant end,case nTipoLeasing when 'L' then 4 else CASE WHEN nGarant = 1 THEN 1          WHEN nGarant = 3 THEN 2          WHEN nGarant = 4 THEN 3 END end "
'  sSql = sSql & " order by TipoGar,TipoCred "
  
  sSql = "EXEC stp_sel_GeneraReporteAnexo2 " & pnAnio & "," & pnMes & "," & pnTpoCambio
'END JUEZ **************************************************************************

    Set rs = oCon.CargaRecordSet(sSql)
    Dim columna As String
    Do While Not rs.EOF
            If rs!TipoCred > 6 Then
               nTipoCred = rs!TipoCred - 1
            Else
               nTipoCred = rs!TipoCred
            End If
                If rs!TipoGar = 1 Then
                    columna = "C"
                ElseIf rs!TipoGar = 2 Then
                    columna = "I"
                ElseIf rs!TipoGar = 3 Then
                    columna = "W"
                ElseIf rs!TipoGar = 4 Then
                    columna = "V"
                End If
                
                xlHoja1.Range(columna & 5 * nTipoCred + 9) = rs!nVigente
                xlHoja1.Range(columna & 5 * nTipoCred + 10) = rs!nRefinanciado
                xlHoja1.Range(columna & 5 * nTipoCred + 11) = rs!nVencido + rs!nJudicial
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
    Loop
    rs.Close
    
    sSql = " select cPlazo, "
    sSql = sSql & " Saldo = SUM(Case When cMoneda = '1' Then nSaldo Else nSaldo*" & pnTpoCambio & " End) "
    sSql = sSql & " from " & sservidorconsolidada & "CredSaldosCont "
    sSql = sSql & " where dFecha = '" & Format(ldFecha, "yyyyMMdd") & "' and cConcepto = 'C' AND cEstado = '1' "
    sSql = sSql & " Group by cPlazo "
    
    Set rs = oCon.CargaRecordSet(sSql)
    Do While Not rs.EOF
        xlHoja1.Range("Y" & 48 + rs!cPlazo) = rs!Saldo
        rs.MoveNext
        If rs.EOF Then
           Exit Do
        End If
    Loop
    rs.Close
    
    oBarra.Progress 2 * 2 - 1, lsTitulo, "Generando reporte...", , vbBlue

    xlHoja1.SaveAs App.path & lsArchivo1
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    
   oCon.CierraConexion
   
End Sub


Private Sub FormatoReporteFinal()

xlHoja1.Range("F1:I1").ColumnWidth = 3
xlHoja1.Range("K1:V1").ColumnWidth = 3
xlHoja1.Range("AD1:AD1").ColumnWidth = 12

xlHoja1.Range("B7:AD11").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B7:AD11").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("B7:AD11").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B12:AD45").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B12:AD45").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B25:AD25").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B25:AD25").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B38:AD38").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B38:AD38").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B12:C45").Font.Size = 8
xlHoja1.Range("B12:C45").Font.Name = "Arial"
xlHoja1.Range("B12:B45").ColumnWidth = 21.15
xlHoja1.Range("C12:C45").ColumnWidth = 24.14

xlHoja1.Range("R12:R45").ColumnWidth = 8

xlHoja1.Range("D52:G52").MergeCells = True
xlHoja1.Range("T52:W52").MergeCells = True
xlHoja1.Range("T53:W53").MergeCells = True
 
xlHoja1.Cells(52, 4) = "Gerente General"
xlHoja1.Cells(52, 20) = "Contador General"
xlHoja1.Cells(53, 20) = "Matricula No"

xlHoja1.Range("D12:AD44").NumberFormat = "###,###"
xlHoja1.Range("D45:AD45").NumberFormat = "###,##0"

xlHoja1.Range("B46:AD45").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("D52:G52").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("T52:W52").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("D52:W53").HorizontalAlignment = xlCenter
xlHoja1.Range("D52:W53").Font.Size = 10
xlHoja1.Range("D52:W53").Font.Name = "Arial"
 
End Sub


Public Sub GeneraAnx05DeudoresProvisiones(pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim n           As Integer
Dim lsNomHoja   As String
On Error GoTo ErrImprime

lnTpoCambio = pnTpoCambio

If pnBandera = 1 Then
    lsTitulo = "ANEXO 5"
    Set oBarra = New clsProgressBar
    oBarra.Max = 1
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionOnly
    
    oBarra.Progress 0, lsTitulo, "Generando Hoja Excel...", , vbBlue
    MousePointer = 11
    
    lsArchivo = App.path & "\Spooler\Anx05_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)) & gsCodUser, Trim(Str(pnMes))) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          For n = 1 To 2
            If n = 1 Then
                lsNomHoja = "En Miles"
            Else
                lsNomHoja = "En Nuevos Soles"
            End If
            ExcelAddHoja lsNomHoja, xlLibro, xlHoja1
            oBarra.Progress 1, lsTitulo, "Hoja Excel generada...", , vbBlue
            Call GeneraReporteAnexo5(pnAnio, pnMes, pnTpoCambio, psMes, n)
          Next
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
       End If
    MousePointer = 0
    oBarra.CloseForm Me
    Set oBarra = Nothing
    MsgBox "Archivo del Anexo5 generado satisfactoriamente", vbInformation, "Aviso!!!"
ElseIf pnBandera = 2 Then
     'GeneraSUCAVEAnx05 False, pnMes, pnAnio, psMes
End If

Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   If Not oBarra Is Nothing Then
      oBarra.CloseForm Me
   End If
   Set oBarra = Nothing
   MousePointer = 0
End Sub


Private Sub AsignaImporteAdministrativo(pnCol As Integer, psMoneda As String, Optional nRow As Integer = 0)
'   sSql = "SELECT sum(nSaldoCap) nSaldo from dbRcd..CredAdmin " _
'        & "WHERE cEstado IN ('F','V','1','4','6','7') and cCodCta LIKE '_____" & psMoneda & "%'"
'   RSClose R
'   Set R = CargaRecord(sSql)
'   If Not R.EOF Then
'      xlHoja1.Range(xlHoja1.Cells(16 + nRow, pnCol), xlHoja1.Cells(16 + nRow, pnCol)).Formula = xlHoja1.Range(xlHoja1.Cells(16 + nRow, pnCol), xlHoja1.Cells(16 + nRow, pnCol)).Formula & "+" & Round(IIf(IsNull(R!nSaldo), 0, R!nSaldo) * IIf(psMoneda = "2", Val(txtTipCambio), 1), 2)
'   End If
'   RSClose R
End Sub

Private Sub AsignaImporteIndirecto(pdFecha As Date)
Dim sSql As String
Dim reg As New ADODB.Recordset
Dim nImporteSoles As Currency
Dim nImporteDolares As Currency

    If gbBitCentral = True Then
        sSql = " SELECT substring(cCtaCod, 9,1) as nMoneda, sum(nSaldoCap) as nImporte " & _
                " FROM CartaFianzaSaldoConsol " & _
                " WHERE CONVERT(VARCHAR(8), dFecha,112)='" & Format(pdFecha, "YYYYMMdd") & "' " & _
                " AND nPrdEstado IN(" & cVigente & ")" & _
                " AND cCtaCod LIKE '_____421[12]%' " & _
                " GROUP BY Substring(cCtaCod, 9,1) "
    Else
        sSql = " SELECT substring(cCodCta, 6,1) as nMoneda, sum(nSaldoCap) as nImporte " & _
                " FROM CartaFianzaSaldoConsol " & _
                " WHERE CONVERT(VARCHAR(8), dFecha,112)='" & Format(pdFecha, "YYYYMMdd") & "' " & _
                " AND cEstado='F' " & _
                " AND cCodCta LIKE '__421[12]%' " & _
                " GROUP BY Substring(cCodCta, 6,1) "
    End If
    
Set reg = oCon.CargaRecordSet(sSql)
Do While Not reg.EOF
    If reg!nMoneda = "1" Then
        nImporteSoles = reg!nImporte
    ElseIf reg!nMoneda = "2" Then
        nImporteDolares = reg!nImporte
    End If
    reg.MoveNext
Loop
reg.Close
Set reg = Nothing
 
    xlHoja1.Cells(41, 30) = nImporteSoles
    xlHoja1.Cells(42, 30) = Round(nImporteDolares * lnTpoCambio, 2)
     
    xlHoja1.Range(xlHoja1.Cells(39, 30), xlHoja1.Cells(39, 30)).Formula = "=AD41+AD42"
    xlHoja1.Range("AF39:AF39").Formula = "=SUM(L39:AE39)"
    xlHoja1.Range("AF41:AF41").Formula = "=SUM(L41:AE41)"
    xlHoja1.Range("AF42:AF42").Formula = "=SUM(L42:AE42)"
    
    xlHoja1.Range("AG39:AG39").Formula = "=K39+AF39"
    xlHoja1.Range("AG41:AG41").Formula = "=K41+AF41"
    xlHoja1.Range("AG42:AG42").Formula = "=K42+AF42"
    
    xlHoja1.Range("AJ39:AJ39").Formula = "=+AG39+AH39+AI39"
    xlHoja1.Range("AJ41:AJ41").Formula = "=+AG41+AH41+AI41"
    xlHoja1.Range("AJ42:AJ42").Formula = "=+AG42+AH42+AI42"
 

End Sub

Private Sub AsignaImporteFila(pnItem As Integer, psMoneda As String, Optional pnRow As Integer = 0, Optional ldFecha As Date)
Select Case pnItem
   Case 1:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 16, cVigente & ", " & cPigno, "1"               'Vig CP
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 16, "'F','1','4','6','7'", "1"            'Vig CP
        End If
   Case 2:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 17, cVigente & ", " & cPigno, "2"         'Vig LP
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 17, "'F','1','4','6','7'", "2"         'Vig LP
        End If
   Case 3:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 19, cVigente & ", " & cPigno, , 0, True       'Refinanc Vigentes
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 19, "'F','1','4','6','7'", , 0, True       'Refinanc Vigentes
        End If
   Case 4:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 21, cVigente & ", " & cPigno, , 1         'Venc <=4
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 21, "'F','1','4','6','7'", , 1         'Venc <=4
        End If
   Case 5:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 22, cVigente & ", " & cPigno, , 2        'Venc > 4
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 22, "'F', '1','4','6','7'", , 2        'Venc > 4
        End If
   Case 6:
        If gbBitCentral = True Then
            GetImporteCredito ldFecha, psMoneda, pnRow + 23, gColocEstRecVigJud, , 3           'Judicial
        Else
            GetImporteCredito ldFecha, psMoneda, pnRow + 23, "'V'", , 3           'Judicial
        End If
End Select
End Sub

Private Sub GetImporteCredito(pdFecha As Date, psMoneda As String, pnRow As Integer, Optional psEstado As String = "", Optional psLinCred As String = "", Optional pnDiasAtraso As Integer = 0, Optional pbRefina As Boolean = False)
Dim pSql As String
Dim prs As ADODB.Recordset
Dim lncol As Integer
 
Dim sCondEsta As String
Dim sCondDias As String
Dim sCondCred As String
Dim sImporte  As String
Dim sJoinCred As String
sCondDias = "": sCondCred = "": sCondEsta = "": sImporte = ""

If psEstado <> "" And Not pnDiasAtraso = 2 Then
    If gbBitCentral = True Then
        sCondEsta = " and csc.nprdEstado IN(" & psEstado & ") "
    Else
        sCondEsta = " and csc.cEstado IN (" & psEstado & ") "
    End If
End If
  
If pnDiasAtraso = 0 Then 'Vigentes
    If gbBitCentral = True Then
        sCondDias = " and ((csc.cCtaCod LIKE '_____1%' and nDiasAtraso < 16) or (csc.cCtaCod LIKE '_____2%' and nDiasAtraso < 31) or (csc.cCtaCod LIKE '_____[34]%' and not csc.cCtaCod LIKE '_____305%' and nDiasAtraso <= 90) or (csc.cCtaCod LIKE '_____305%' and nDiasAtraso < 31) ) "
        sImporte = " ISNULL(SUM(CASE WHEN csc.cCtaCod LIKE '_____[34]%' and not csc.cCtaCod LIKE '_____305%' and nDiasAtraso >=31 and nDiasAtraso <=90 THEN csc.nSaldoCap - nCapVencido "
        sImporte = sImporte & "                 ELSE csc.nSaldoCap END ),0) "
    Else
        sCondDias = " and ((csc.cCodCta LIKE '__1%' and nDiasAtraso < 16) or (csc.cCodCta LIKE '__2%' and nDiasAtraso < 31) or (csc.cCodCta LIKE '__[34]%' and not csc.cCodCta LIKE '__305%' and nDiasAtraso <= 90) or (csc.cCodCta LIKE '__305%' and nDiasAtraso < 31) ) "
        sImporte = " ISNULL(SUM(CASE WHEN csc.cCodCta LIKE '__[34]%' and not csc.cCodCta LIKE '__305%' and nDiasAtraso >=31 and nDiasAtraso <=90 THEN csc.nSaldoCap - nCapVencido " _
                 & "                 ELSE csc.nSaldoCap END ),0) "
    End If
ElseIf pnDiasAtraso = 1 Then 'Vencidos < 120 dias
    If gbBitCentral = True Then
        sCondDias = " and ((csc.cCtaCod LIKE '_____1%' and nDiasAtraso >= 16 and nDiasAtraso <= 120) or (csc.cCtaCod LIKE '_____[234]%' and nDiasAtraso >= 31 and nDiasAtraso <= 120)) "
        sImporte = " ISNULL(SUM(CASE WHEN csc.cCtaCod LIKE '_____[34]%' and not csc.cCtaCod LIKE '_____305%' and nDiasAtraso >=31 and nDiasAtraso <=90 THEN nCapVencido "
        sImporte = sImporte & "                 WHEN csc.cCtaCod LIKE '_____[34]%' and not csc.cCtaCod LIKE '_____305%' and nDiasAtraso > 90 THEN csc.nSaldoCap "
        sImporte = sImporte & "                 ELSE csc.nSaldoCap END ),0) "
    Else
        sCondDias = " and ((csc.cCodCta LIKE '__1%' and nDiasAtraso >= 16 and nDiasAtraso <= 120) or (csc.cCodCta LIKE '__[234]%' and nDiasAtraso >= 31 and nDiasAtraso <= 120)) "
        sImporte = " ISNULL(SUM(CASE WHEN csc.cCodCta LIKE '__[34]%' and not csc.cCodCta LIKE '__305%' and nDiasAtraso >=31 and nDiasAtraso <=90 THEN nCapVencido " _
                & "                 WHEN csc.cCodCta LIKE '__[34]%' and not csc.cCodCta LIKE '__305%' and nDiasAtraso > 90 THEN csc.nSaldoCap " _
                & "                 ELSE csc.nSaldoCap END ),0) "
    End If
ElseIf pnDiasAtraso = 2 Then 'Vencidos < 120 dias
    If gbBitCentral = True Then
        sJoinCred = " JOIN CreditoConsol c ON c.cCtaCod = csc.cCtaCod and c.nPrdEstado = csc.nPrdEstado "
        sCondDias = " and ((csc.nPrdEstado IN (" & cVigente & ", " & cPigno & ") and csc.nDiasAtraso > 120)  or (csc.nPrdEstado=" & gColocEstRecVigJud & " and nDemanda = " & gColRecDemandaNo & "))  "
        sImporte = " ISNULL(SUM(csc.nSaldoCap),0)"
    Else
        sJoinCred = " JOIN CreditoConsol c ON c.cCodCta = csc.cCodCta and c.cEstado = csc.cEstado "
        sCondDias = " and ((csc.cEstado IN ('F','1','4','6','7') and csc.nDiasAtraso > 120)  or (csc.cEstado = 'V' and c.cCondCre = 'J' and cDemanda = 'N'))  "
        sImporte = " ISNULL(SUM(csc.nSaldoCap),0)"
    End If
Else
    If gbBitCentral = True Then
        sJoinCred = "JOIN CreditoConsol c ON c.cCtaCod = csc.cCtaCod and c.nDemanda=" & gColRecDemandaSi & " "
        sCondDias = ""
        sImporte = " ISNULL(SUM(csc.nSaldoCap),0)"
    Else
        sJoinCred = "JOIN CreditoConsol c ON c.cCodCta = csc.cCodCta and c.cDemanda = 'S' "
        sCondDias = ""
        sImporte = " ISNULL(SUM(csc.nSaldoCap),0)"
    End If
End If

'Refinanciado y Linea de Credito
If pbRefina And psLinCred = "" Then
    If gbBitCentral = True Then
        sCondCred = " and csc.cCtaCod IN (SELECT DISTINCT cCtaCod FROM CreditoConsol csc WHERE cRefinan='R') "
    Else
        sCondCred = " and csc.cCodCta IN (SELECT DISTINCT cCodCta FROM CreditoConsol csc WHERE cRefinan = 'R') "
    End If
End If
If Not pbRefina And psLinCred <> "" Then
    If gbBitCentral = True Then
        sCondCred = " and csc.cCtaCod IN (SELECT DISTINCT cCtaCod FROM CreditoConsol csc WHERE nPrdEstado IN (" & cVigente & ", " & cPigno & ") and cLineaCred like '_____" & psLinCred & "%' and cRefinan <>'R') "
        If psLinCred = "2" Then
           sCondCred = sCondCred & " and not csc.cCtaCod IN (SELECT DISTINCT cCtaCod FROM CreditoConsol csc WHERE nPrdEstado IN (" & cVigente & ", " & cPigno & ") and cLineaCred like '_____1%' and cRefinan <>'R') "
        End If
    Else
        sCondCred = " and csc.cCodCta IN (SELECT DISTINCT cCodCta FROM CreditoConsol csc WHERE cEstado IN ('F','1','4','6','7') and cCodLinCred like '____" & psLinCred & "%' and cRefinan <> 'R') "
        If psLinCred = "2" Then
           sCondCred = sCondCred & " and not csc.cCodCta IN (SELECT DISTINCT cCodCta FROM CreditoConsol csc WHERE cEstado IN ('F','1','4','6','7') and cCodLinCred like '____1%' and cRefinan <> 'R') "
        End If
    End If
End If
If pbRefina And psLinCred <> "" Then
    If gbBitCentral = True Then
        sCondCred = " and csc.cCtaCod IN (SELECT DISTINCT cCtaCod FROM CreditoConsol csc WHERE cLineaCred like '_____" & psLinCred & "%' and cRefinan ='R') "
        If psLinCred = "2" Then
           sCondCred = sCondCred & " and not csc.cCtaCod IN (SELECT DISTINCT cCtaCod FROM CreditoConsol csc WHERE cLineaCred like '_____1%' and cRefinan ='R' "
        End If
    Else
        sCondCred = " and csc.cCodCta IN (SELECT DISTINCT cCodCta FROM CreditoConsol csc WHERE cCodLinCred like '____" & psLinCred & "%' and cRefinan = 'R') "
        If psLinCred = "2" Then
           sCondCred = sCondCred & " and not csc.cCodCta IN (SELECT DISTINCT cCodCta FROM CreditoConsol csc WHERE cCodLinCred like '____1%' and cRefinan = 'R') "
        End If
    End If
End If

If gbBitCentral = True Then
    pSql = "SELECT CASE WHEN csc.cCtaCod LIKE '_____305%' THEN '07' ELSE ISNULL(gc1.cTipoGarant,'12') END as cTipoGarant, " & sImporte & " AS nImporte " _
        & "FROM CreditoSaldoConsol csc " & sJoinCred & " LEFT JOIN " _
        & "    ( SELECT gcc.cCtaCod, MIN(gc.cTipoGarant) cTipoGarant " _
        & "      FROM GarantCredConsol gcc JOIN GarantiasConsol gc ON gc.cNumGarant = gcc.cNumGarant " _
        & "      GROUP BY gcc.cCtaCod " _
        & "    ) as gc1 ON gc1.cCtaCod = csc.cCtaCod " _
        & "WHERE csc.cCtaCod LIKE '________" & psMoneda & "%' and " _
        & "      datediff(dd,csc.dfecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 " & sCondEsta _
        & sCondDias & sCondCred _
        & "GROUP BY CASE WHEN csc.cCtaCod LIKE '_____305%' THEN '07' ELSE ISNULL(gc1.cTipoGarant,'12') END "
Else
    pSql = "SELECT CASE WHEN csc.cCodCta LIKE '__305%' THEN '07' ELSE ISNULL(gc1.cTipoGarant,'12') END as cTipoGarant, " & sImporte & " AS nImporte " _
         & "FROM CreditoSaldoConsol csc " & sJoinCred & " LEFT JOIN " _
         & "    ( SELECT gcc.cCodCta, MIN(gc.cTipoGarant) cTipoGarant " _
         & "      FROM GarantCredConsol gcc JOIN GarantiasConsol gc ON gc.cNumGarant = gcc.cNumGarant " _
         & "      GROUP BY gcc.cCodCta " _
         & "    ) as gc1 ON gc1.cCodCta = csc.cCodCta " _
         & "WHERE csc.cCodCta LIKE '_____" & psMoneda & "%' and " _
         & "      datediff(dd,csc.dfecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 " & sCondEsta _
         & sCondDias & sCondCred _
         & "GROUP BY CASE WHEN csc.cCodCta LIKE '__305%' THEN '07' ELSE ISNULL(gc1.cTipoGarant,'12') END "
End If

Set prs = oCon.CargaRecordSet(pSql, adLockReadOnly)
Do While Not prs.EOF
   Select Case Format(prs!cTipoGarant, "00")
      Case "01"
            lncol = 12
      Case "02"
            lncol = 23
      Case "03"
            lncol = 22
      Case "04"
            lncol = 25
      Case "05"
            lncol = 34
      Case "06"
            lncol = 4
      Case "07"
            lncol = 20
      Case "08"
            lncol = 26
      Case "09"
            lncol = 34
      Case "10"
            lncol = 34
      Case "11"
            lncol = 34
      Case "12"
            lncol = 34
   End Select
   
    If xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "" Then
          xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "="
    End If
    
    xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula & "+" & Round(prs!nImporte * IIf(psMoneda = "2", lnTpoCambio, 1), 2)
    prs.MoveNext
Loop
RSClose prs
End Sub

Public Sub GeneraAnx02CreditosTipoGarantia(pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim n           As Integer

On Error GoTo ErrImprime

lnTpoCambio = pnTpoCambio

If pnBandera = 1 Then
    lsTitulo = "CREDITOS POR TIPO DE GARANTIA"
    Set oBarra = New clsProgressBar
    oBarra.Max = 1
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionOnly
    
    oBarra.Progress 0, lsTitulo, "Generando Hoja Excel...", , vbBlue
    MousePointer = 11
    
    lsArchivo = App.path & "\Spooler\Anx02_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)) & gsCodUser, Trim(Str(pnMes))) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          ExcelAddHoja psMes, xlLibro, xlHoja1
          oBarra.Progress 1, lsTitulo, "Hoja Excel generada...", , vbBlue
          Call GeneraReporteAnexo2(pnAnio, pnMes, pnTpoCambio, psMes)
          'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
          'CargaArchivo lsArchivo, App.path & "\Spooler"
       End If
    MousePointer = 0
    oBarra.CloseForm Me
    Set oBarra = Nothing
    MsgBox "Archivo generado satisfactoriamente", vbInformation, "Aviso!!!"
ElseIf pnBandera = 2 Then
     GeneraSUCAVEAnx02 True, pnMes, pnAnio, psMes
     
End If

Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   If Not oBarra Is Nothing Then
      oBarra.CloseForm Me
   End If
   Set oBarra = Nothing
   MousePointer = 0
End Sub

Private Sub CabeceraExcelAnexo2(pdFecha As Date, psMes As String)

xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterVertically = True
xlHoja1.PageSetup.Zoom = 50
xlHoja1.Cells(1, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
xlHoja1.Cells(2, 2) = "EMPRESA : " & gsNomCmac
xlHoja1.Cells(2, 25) = "CODIGO : " & gsCodCMAC
xlHoja1.Cells(2, 29) = "ANEXO NRO. 02"
xlHoja1.Cells(3, 2) = " CREDITOS DIRECTOS POR TIPO DE GARANTIA "
xlHoja1.Cells(4, 2) = "Al " & Mid(pdFecha, 1, 2) & " de " & Trim(psMes) & " de " & Year(pdFecha)
xlHoja1.Cells(5, 2) = " "

xlHoja1.Range("B3:AD3").Merge
xlHoja1.Range("B4:AD4").Merge
xlHoja1.Range("B5:AD5").Merge

xlHoja1.Range("B3:AD3").HorizontalAlignment = xlCenter
xlHoja1.Range("B4:AD4").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:AD5").HorizontalAlignment = xlCenter


xlHoja1.Cells(8, 3) = "DENOMINACION"
xlHoja1.Range("C8:C11").Merge
xlHoja1.Range("C8:C11").VerticalAlignment = xlCenter
xlHoja1.Range("C8:C11").HorizontalAlignment = xlCenter
xlHoja1.Range("C8:C11").WrapText = True

xlHoja1.Cells(7, 4) = " Créditos con garantía de depósito en efectivo  -1 "
xlHoja1.Range("D8:D11").Font.Size = 7
xlHoja1.Range("D8:D11").Merge
xlHoja1.Range("D8:D11").VerticalAlignment = xlCenter
xlHoja1.Range("D8:D11").HorizontalAlignment = xlCenter
xlHoja1.Range("D8:D11").WrapText = True

xlHoja1.Cells(7, 5) = "  Créditos con garantía de derechos de carta de crédito irrevocables  -2  "
xlHoja1.Range("E8:E11").Font.Size = 7
xlHoja1.Range("E8:E11").Merge
xlHoja1.Range("E8:E11").VerticalAlignment = xlCenter
xlHoja1.Range("E8:E11").HorizontalAlignment = xlCenter
xlHoja1.Range("E8:E11").WrapText = True

xlHoja1.Cells(7, 6) = "  Créditos con primera prenda o fideicomiso en garantía sobre los siguientes Bienes: "
xlHoja1.Range("F8:I10").Font.Size = 7
xlHoja1.Range("F8:I10").Merge
xlHoja1.Range("F8:I10").VerticalAlignment = xlCenter
xlHoja1.Range("F8:I10").HorizontalAlignment = xlCenter
xlHoja1.Range("F8:I10").WrapText = True
xlHoja1.Cells(11, 7) = "(3)"
xlHoja1.Cells(11, 8) = "(4)"
xlHoja1.Cells(11, 9) = "(5)"
xlHoja1.Cells(11, 10) = "(6)"
xlHoja1.Range("F11:I12").HorizontalAlignment = xlCenter

xlHoja1.Cells(7, 10) = "  Créditos con primera hipoteca o con fiducia en garantía sobre inmuebles"
xlHoja1.Range("J7:J11").Font.Size = 7
xlHoja1.Range("J7:J11").Merge
xlHoja1.Range("J7:J11").VerticalAlignment = xlCenter
xlHoja1.Range("J7:J11").HorizontalAlignment = xlCenter
xlHoja1.Range("J7:J11").WrapText = True

xlHoja1.Cells(7, 11) = " Créditos con primera prenda o con fiducia en garantía  sobre los siguientes bienes      "
xlHoja1.Range("K7:V11").Font.Size = 7
xlHoja1.Range("K7:V10").Merge
xlHoja1.Range("K7:V10").HorizontalAlignment = xlCenter
xlHoja1.Range("K7:V10").VerticalAlignment = xlCenter
xlHoja1.Range("K7:V10").WrapText = True
xlHoja1.Cells(11, 11) = "(7)"
xlHoja1.Cells(11, 12) = "(8)"
xlHoja1.Cells(11, 13) = "(9)"
xlHoja1.Cells(11, 14) = "(10)"
xlHoja1.Cells(11, 15) = "(11)"
xlHoja1.Cells(11, 16) = "(12)"
xlHoja1.Cells(11, 17) = "(13)"
xlHoja1.Cells(11, 18) = "(14)"
xlHoja1.Cells(11, 19) = "(15)"
xlHoja1.Cells(11, 20) = "(16)"
xlHoja1.Cells(11, 21) = "(17)"
xlHoja1.Cells(11, 22) = "(18)"
xlHoja1.Range("K11:V11").HorizontalAlignment = xlCenter

xlHoja1.Cells(7, 23) = "  Créditos con primera prenda agrícola o minera Sobre bienes de fácil realización"
xlHoja1.Range("W7:W11").Font.Size = 7
xlHoja1.Range("W7:W11").Merge
xlHoja1.Range("W7:W11").VerticalAlignment = xlCenter
xlHoja1.Range("W7:W11").HorizontalAlignment = xlCenter
xlHoja1.Range("W7:W11").WrapText = True

xlHoja1.Cells(7, 24) = "Créditos con primera prenda global y flotante "
xlHoja1.Range("X7:X11").Font.Size = 7
xlHoja1.Range("X7:X11").Merge
xlHoja1.Range("X7:X11").VerticalAlignment = xlCenter
xlHoja1.Range("X7:X11").HorizontalAlignment = xlCenter
xlHoja1.Range("X7:X11").WrapText = True

xlHoja1.Cells(7, 25) = " Créditos con seguro de crédito a la exportación para el financiamiento pre y post-embarque cubiertos por la póliza respectiva  "
xlHoja1.Range("Y7:Y11").Font.Size = 7
xlHoja1.Range("Y7:Y11").Merge
xlHoja1.Range("Y7:Y11").VerticalAlignment = xlCenter
xlHoja1.Range("Y7:Y11").HorizontalAlignment = xlCenter
xlHoja1.Range("Y7:Y11").WrapText = True

xlHoja1.Cells(7, 26) = "Créditos con póliza del programa de seguro de crédito para la pequeña empresa    (19) "
xlHoja1.Range("Z7:Z11").Font.Size = 7
xlHoja1.Range("Z7:Z11").Merge
xlHoja1.Range("Z7:Z11").VerticalAlignment = xlCenter
xlHoja1.Range("Z7:Z11").HorizontalAlignment = xlCenter
xlHoja1.Range("Z7:Z11").WrapText = True

xlHoja1.Cells(7, 27) = "   Créditos con pólizas de caución (20) "
xlHoja1.Range("AA7:AA11").Font.Size = 7
xlHoja1.Range("AA7:AA11").Merge
xlHoja1.Range("AA7:AA11").VerticalAlignment = xlCenter
xlHoja1.Range("AA7:AA11").HorizontalAlignment = xlCenter
xlHoja1.Range("AA7:AA11").WrapText = True

xlHoja1.Cells(7, 28) = "  Créditos con garantías no preferidas "
xlHoja1.Range("AB7:AB11").Font.Size = 7
xlHoja1.Range("AB7:AB11").Merge
xlHoja1.Range("AB7:AB11").VerticalAlignment = xlCenter
xlHoja1.Range("AB7:AB11").HorizontalAlignment = xlCenter
xlHoja1.Range("AB7:AB11").WrapText = True

xlHoja1.Cells(7, 29) = " Créditos sin garantías "
xlHoja1.Range("AC7:AC11").Font.Size = 7
xlHoja1.Range("AC7:AC11").Merge
xlHoja1.Range("AC7:AC11").VerticalAlignment = xlCenter
xlHoja1.Range("AC7:AC11").HorizontalAlignment = xlCenter
xlHoja1.Range("AC7:AC11").WrapText = True

xlHoja1.Cells(7, 30) = " Total Créditos "
xlHoja1.Range("AD7:AD11").Font.Size = 7
xlHoja1.Range("AD7:AD11").Merge
xlHoja1.Range("AD7:AD11").VerticalAlignment = xlCenter
xlHoja1.Range("AD7:AD11").HorizontalAlignment = xlCenter
xlHoja1.Range("AD7:AD11").WrapText = True

xlHoja1.Range("A1:A1").ColumnWidth = 1
xlHoja1.Range("B1:B1").ColumnWidth = 9
xlHoja1.Range("C1:C1").ColumnWidth = 30
xlHoja1.Range("D1:D1").ColumnWidth = 10
xlHoja1.Range("E1:E1").ColumnWidth = 12
xlHoja1.Range("J1:J1").ColumnWidth = 12
xlHoja1.Range("W1:W1").ColumnWidth = 12
xlHoja1.Range("X1:X1").ColumnWidth = 12
xlHoja1.Range("Y1:Y1").ColumnWidth = 15
xlHoja1.Range("Z1:Z1").ColumnWidth = 13
xlHoja1.Range("AA1:AA1").ColumnWidth = 12
xlHoja1.Range("AB1:AB1").ColumnWidth = 12
xlHoja1.Range("AC1:AC1").ColumnWidth = 12
xlHoja1.Range("AD1:AD1").ColumnWidth = 14

xlHoja1.Range("D12:AD44").NumberFormat = "##,###;-##,###"
xlHoja1.Range("D45:AD45").NumberFormat = "##,##0;-##,##0"

xlHoja1.Range("B1:C1").EntireColumn.HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("B1:C2").HorizontalAlignment = xlHAlignLeft
End Sub



Private Sub GetFileAnexos(pdFecha As Date)
Dim pSql As String
Dim prs As ADODB.Recordset
Dim lncol As Integer
Dim pnRow As Integer
Dim Monto As Double

pSql = "SELECT nANEXO, nAÑO, nMES, nLINEA, nCOLUMNA, nMONTO  " _
        & " FROM Anexos " _
        & "WHERE nANEXO = 2 AND nSUBAN = 1 AND nAÑO = " & Year(pdFecha) & " AND nMES = " & Month(pdFecha)

Set prs = oCon.CargaRecordSet(pSql, adLockReadOnly)
Do While Not prs.EOF
    pnRow = prs!nlinea
    lncol = prs!nColumna
    'Monto = Int(Round(prs!nMonto / 1000, 0))
    Monto = Format(prs!nMonto, "#0.00")
    If xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "" Then
          xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "="
    End If
    
    xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula & "+" & Monto
    prs.MoveNext
Loop
RSClose prs
End Sub


Private Sub Form_Initialize()
Dim oConst As NConstSistemas

Set oConst = New NConstSistemas

sservidorconsolidada = oConst.LeeConstSistema(gConstSistServCentralRiesgos)

Set oConst = Nothing
End Sub

Private Sub Form_Load()
CentraForm Me

cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "'"
If gsCodCMAC = "102" Then
    cPigno = "'" & gPigEstDesemb & "', '" & gPigEstAmortiz & "', '" & gPigEstReusoLin & "', '" & gPigEstRemat & "'"
Else
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
End If
Set oCon = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oCon = Nothing
End Sub

'''''Public Sub GeneraAnx03FlujoCrediticioPorTipoCred(ByVal pnBitCentral As Boolean, pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
'''''Dim nCol  As Integer
'''''Dim sCol  As String
'''''
'''''Dim lsArchivo   As String
'''''Dim lbLibroOpen As Boolean
'''''Dim N           As Integer
'''''Dim ldFechaRep As Date
'''''
'''''On Error GoTo ErrImprime
'''''
'''''MousePointer = 11
'''''lnTpoCambio = pnTpoCambio
'''''
'''''If pnBandera = 1 Then
'''''    lsArchivo = App.path & "\Spooler\Anx03_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)) & gsCodUser, Trim(Str(pnMes))) & ".xls"
'''''    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
'''''    If lbLibroOpen Then
'''''        ExcelAddHoja psMes, xlLibro, xlHoja1
'''''
'''''        'If pnBitCentral = True Then
'''''        '    Call GeneraReporteAnexo3C(pnMes, pnAnio, pnTpoCambio, psMes)
'''''        'Else
'''''            ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
'''''            Call GeneraReporteAnexo3(ldFechaRep, pnTpoCambio, psMes)
'''''        'End If
'''''
'''''        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'''''        CargaArchivo lsArchivo, App.path & "\Spooler"
'''''    End If
'''''    MousePointer = 0
'''''    MsgBox "Reporte Generado Satisfactoriamente"
'''''ElseIf pnBandera = 2 Then
'''''    MousePointer = 11
'''''
'''''    If pnBitCentral = True Then
'''''    Else
'''''        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
'''''        GeneraSUCAVEAnx03 pnBitCentral, ldFechaRep, psMes
'''''    End If
'''''
'''''End If
'''''
'''''Exit Sub
'''''ErrImprime:
'''''   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
'''''   If lbLibroOpen Then
'''''      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
'''''      lbLibroOpen = False
'''''   End If
'''''   MousePointer = 0
'''''End Sub

Public Sub GeneraAnx03FlujoCrediticioPorTipoCred(ByVal pnBitCentral As Boolean, pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim n           As Integer
Dim ldFechaRep As Date
 
On Error GoTo ErrImprime
 
'MousePointer = 11
lnTpoCambio = pnTpoCambio
 
If pnBandera = 1 Then
    lsArchivo = App.path & "\Spooler\Anx03_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)) & gsCodUser, Trim(Str(pnMes))) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        ExcelAddHoja psMes, xlLibro, xlHoja1
                    
        'If pnBitCentral = True Then
        '    Call GeneraReporteAnexo3C(pnMes, pnAnio, pnTpoCambio, psMes)
        'Else
            ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
            Call GeneraReporteAnexo3(ldFechaRep, pnTpoCambio, psMes)
        'End If
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        CargaArchivo lsArchivo, App.path & "\Spooler"
    End If
    MousePointer = 0
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
ElseIf pnBandera = 2 Then
    MousePointer = 11
    
    If pnBitCentral = True Then
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        GeneraSUCAVEAnx03 pnBitCentral, ldFechaRep, psMes
    Else
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        GeneraSUCAVEAnx03 pnBitCentral, ldFechaRep, psMes
    End If

End If

Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   MousePointer = 0
End Sub


Private Sub GeneraSUCAVEAnx02(pnBitCentral As Boolean, pnMes As Integer, pnannio As Integer, psMes As String)
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim nFil As Integer
Dim nCol As Integer
Dim pdFecha As String
Dim I As Integer
Dim j As Integer
Dim sCad As String
Dim matTImprimir() As Currency

On Error GoTo ErrBegin

psArchivoALeer = App.path & "\Spooler\Anx02_" & pnannio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)), Trim(Str(pnMes))) & gsCodUser & ".xls"

pdFecha = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnannio, "0000"))) - 1

bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If


    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".102"

    Set xlAplicacion = New Excel.Application

    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    '''''''''''''''''''''''''''''
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase(psMes) Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    'Set xlHoja = xlAplicacion.Worksheets(1)

    Dim matArreglo(1 To 27, 1 To 37) As Currency
    Dim nTemp1 As Integer
    Dim nTemp2 As Integer
    
 
    
    For I = 1 To 37
        If I >= 3 And I <= 10 Then
            nTemp1 = I
            nTemp2 = I + 1
        ElseIf I = 12 Then
            nTemp1 = I
            nTemp2 = I
        ElseIf I >= 14 And I <= 37 Then
            nTemp1 = I
            nTemp2 = I - 1
        Else
            nTemp1 = 0
            nTemp2 = 0
        End If
        
        If nTemp1 > 0 Then
'            matArreglo(1, nTemp1) = xlHoja.Cells(16, nTemp2)
'
'            matArreglo(3, nTemp1) = xlHoja.Cells(15, nTemp2)
'            matArreglo(4, nTemp1) = xlHoja.Cells(18, nTemp2)
'            matArreglo(5, nTemp1) = xlHoja.Cells(19, nTemp2)
'            matArreglo(6, nTemp1) = xlHoja.Cells(21, nTemp2)
'            matArreglo(7, nTemp1) = xlHoja.Cells(22, nTemp2)
'
'            matArreglo(2, nTemp1) = xlHoja.Cells(17, nTemp2)
'
'            matArreglo(8, nTemp1) = xlHoja.Cells(20, nTemp2)
'            matArreglo(9, nTemp1) = xlHoja.Cells(23, nTemp2)
'            matArreglo(11, nTemp1) = xlHoja.Cells(24, nTemp2)
'            matArreglo(12, nTemp1) = xlHoja.Cells(29, nTemp2)
'            matArreglo(13, nTemp1) = xlHoja.Cells(30, nTemp2)
'            matArreglo(14, nTemp1) = xlHoja.Cells(28, nTemp2)
'            matArreglo(15, nTemp1) = xlHoja.Cells(31, nTemp2)
'            matArreglo(16, nTemp1) = xlHoja.Cells(32, nTemp2)
'            matArreglo(17, nTemp1) = xlHoja.Cells(34, nTemp2)
'            matArreglo(18, nTemp1) = xlHoja.Cells(35, nTemp2)
'            matArreglo(19, nTemp1) = xlHoja.Cells(33, nTemp2)
'            matArreglo(20, nTemp1) = xlHoja.Cells(36, nTemp2)
'            matArreglo(22, nTemp1) = xlHoja.Cells(37, nTemp2)
'            matArreglo(23, nTemp1) = matArreglo(11, nTemp1) + matArreglo(22, nTemp1)
'            matArreglo(24, nTemp1) = xlHoja.Cells(41, nTemp2)
'            matArreglo(25, nTemp1) = xlHoja.Cells(42, nTemp2)
'            matArreglo(26, nTemp1) = matArreglo(24, nTemp1) + matArreglo(25, nTemp1)
'            matArreglo(27, nTemp1) = xlHoja.Cells(44, nTemp2)
            
            
            
            matArreglo(1, nTemp1) = xlHoja.Cells(16, nTemp2)
            
            matArreglo(3, nTemp1) = xlHoja.Cells(15, nTemp2)
            matArreglo(4, nTemp1) = xlHoja.Cells(18, nTemp2)
            matArreglo(5, nTemp1) = xlHoja.Cells(19, nTemp2)
            matArreglo(6, nTemp1) = xlHoja.Cells(21, nTemp2)
            matArreglo(7, nTemp1) = xlHoja.Cells(22, nTemp2)
            
            matArreglo(2, nTemp1) = xlHoja.Cells(17, nTemp2)
            
            matArreglo(8, nTemp1) = xlHoja.Cells(20, nTemp2)
            matArreglo(9, nTemp1) = xlHoja.Cells(23, nTemp2)
            
            matArreglo(10, nTemp1) = xlHoja.Cells(24, nTemp2)
            matArreglo(11, nTemp1) = xlHoja.Cells(25, nTemp2)
            
            matArreglo(12, nTemp1) = xlHoja.Cells(30, nTemp2)
            matArreglo(14, nTemp1) = xlHoja.Cells(29, nTemp2)
            matArreglo(13, nTemp1) = xlHoja.Cells(31, nTemp2)
            matArreglo(15, nTemp1) = xlHoja.Cells(32, nTemp2)
            matArreglo(17, nTemp1) = xlHoja.Cells(35, nTemp2)
            matArreglo(18, nTemp1) = xlHoja.Cells(36, nTemp2)
            matArreglo(16, nTemp1) = xlHoja.Cells(33, nTemp2)
            
            
            matArreglo(19, nTemp1) = xlHoja.Cells(34, nTemp2)
            
             
            
            matArreglo(20, nTemp1) = xlHoja.Cells(37, nTemp2)
            
            matArreglo(21, nTemp1) = xlHoja.Cells(38, nTemp2)
            matArreglo(22, nTemp1) = xlHoja.Cells(39, nTemp2)
            matArreglo(23, nTemp1) = xlHoja.Cells(25, nTemp2) + xlHoja.Cells(39, nTemp2)
            
            
            matArreglo(24, nTemp1) = xlHoja.Cells(43, nTemp2)
             
            matArreglo(25, nTemp1) = xlHoja.Cells(44, nTemp2)
            matArreglo(26, nTemp1) = xlHoja.Cells(43, nTemp2) + xlHoja.Cells(44, nTemp2)
            
            matArreglo(27, nTemp1) = xlHoja.Cells(46, nTemp2)
            
            
        End If
    Next
    
    'Calculo los que no aparecen
    
    For I = 1 To 27
        matArreglo(I, 2) = matArreglo(I, 3) + matArreglo(I, 4) + matArreglo(I, 5) + matArreglo(I, 6) + matArreglo(I, 7)
        matArreglo(I, 13) = matArreglo(I, 14) + matArreglo(I, 15) + matArreglo(I, 16) + matArreglo(I, 17) + matArreglo(I, 18) + matArreglo(I, 19) + matArreglo(I, 20) + matArreglo(I, 21) + matArreglo(I, 22) + matArreglo(I, 23) + matArreglo(I, 24) + matArreglo(I, 25)
    Next
    
     
    Open psArchivoAGrabar For Output As #1

    Print #1, "01020100" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012" '0& "000000000000000"
      
    sCad = ""
    
    For I = 1 To 37
        sCad = ""
        For j = 1 To 27
            sCad = sCad & LlenaCerosSUCAVE(matArreglo(j, I))
        Next
        Print #1, IIf(I * 10 < 100, "  ", " ") & Trim(Str(I * 10)) & sCad
    Next
    
    Close #1

    ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False

    MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
    Exit Sub

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
'''''
'''''Private Sub GeneraSUCAVEAnx03(pnBitCentral As Boolean, pdFecha As Date, psMes As String)
'''''Dim psArchivoA_Leer As String
'''''Dim psArchivoAGrabar As String
'''''Dim xlAplicacion As Excel.Application
'''''Dim xlLibro As Excel.Workbook
'''''Dim xlHoja As Excel.Worksheet
'''''Dim bExiste As Boolean
'''''Dim bEncontrado As Boolean
'''''Dim fs As New Scripting.FileSystemObject
'''''Dim matTImprimir(35, 20) As Currency
'''''Dim sCad As String
'''''Dim I As Integer
'''''Dim J As Integer
'''''Dim nFil As Integer
'''''Dim nCol As Integer
'''''
'''''On Error GoTo ErrBegin
'''''
''''''psArchivoA_Leer = App.path & "\Spooler\ANX3_" & Format(gdFecSis & " " & Time, "yyyymmdd") & gsCodUser & ".xls"
'''''
'''''psArchivoA_Leer = App.path & "\Spooler\Anx03_" & Mid(pdFecha, 7, 4) & Mid(pdFecha, 4, 2) & gsCodUser & ".xls"
'''''
'''''bExiste = fs.FileExists(psArchivoA_Leer)
'''''
'''''If bExiste = False Then
'''''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoA_Leer, vbExclamation, "Aviso!!!"
'''''    Exit Sub
'''''End If
'''''
'''''    'Anexo 01 del 15A
'''''    '================
'''''
'''''    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".103"
'''''
'''''    Set xlAplicacion = New Excel.Application
'''''
'''''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoA_Leer)
'''''    '''''''''''''''''''''''''''''
'''''    bEncontrado = False
'''''    For Each xlHoja In xlLibro.Worksheets
'''''        If UCase(xlHoja.Name) = UCase(psMes) Then
'''''            bEncontrado = True
'''''            xlHoja.Activate
'''''            Exit For
'''''        End If
'''''    Next
'''''
'''''    If bEncontrado = False Then
'''''        ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja, True
'''''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
'''''        Exit Sub
'''''    End If
'''''    '''''''''''''''''''''''''''''
'''''    'Set xlHoja = xlAplicacion.Worksheets(1)
'''''
'''''
'''''
'''''    For I = 1 To 35
'''''        If I = 1 Then
'''''            nFil = I + 11
'''''        ElseIf I >= 2 And I <= 4 Then
'''''            nFil = I + 11
'''''        ElseIf I = 5 Then 'Aca se deben sumar desede Alimentos, Bebidas y Tabaco hasta RestoManufactura
'''''        ElseIf I >= 6 And I <= 15 Then
'''''            nFil = I + 10
'''''        ElseIf I >= 16 And I <= 17 Then
'''''            nFil = I + 10
'''''        ElseIf I = 18 Then 'Aca se deben sumar desde Venta y Reparacion de Vehiculos hasta comercio al por menor
'''''        ElseIf I >= 19 And I <= 21 Then
'''''            nFil = I + 9
'''''        ElseIf I >= 22 And I <= 24 Then
'''''            nFil = I + 9
'''''        ElseIf I = 25 Then 'Aca se deben sumar desde Actividad Inmoviliaria y de Alquiler hasta actividad empresarial
'''''        ElseIf I >= 26 And I <= 27 Then
'''''            nFil = I + 8
'''''        ElseIf I >= 28 And I <= 32 Then
'''''            nFil = I + 8
'''''        ElseIf I = 33 Then
'''''            nFil = 44
'''''        ElseIf I = 34 Then
'''''            nFil = 42
'''''        ElseIf I = 35 Then
'''''            nFil = 48
'''''        End If
'''''
'''''        If I <> 5 And I <> 18 And I <> 25 Then
'''''            For J = 2 To 20
'''''                If J = 2 Then
'''''                    If nFil <> 12 And nFil <> 42 And nFil <> 44 And nFil <> 48 Then
'''''                        nCol = J
'''''                        matTImprimir(I, J) = "1" & Mid(Trim(xlHoja.Cells(nFil, nCol)), 1, 2) & Mid(Trim(xlHoja.Cells(nFil, nCol)), 6, 2)
'''''                    End If
'''''                Else
'''''                    nCol = J
'''''                    matTImprimir(I, J) = xlHoja.Cells(nFil, nCol)
'''''                End If
'''''            Next
'''''        End If
'''''    Next
'''''
'''''    For I = 3 To 20
'''''        matTImprimir(5, I) = matTImprimir(6, I) + matTImprimir(7, I) + matTImprimir(8, I) + matTImprimir(9, I) + matTImprimir(10, I) + matTImprimir(11, I) + matTImprimir(12, I) + matTImprimir(13, I) + matTImprimir(14, I) + matTImprimir(15, I)
'''''        matTImprimir(18, I) = matTImprimir(19, I) + matTImprimir(20, I) + matTImprimir(21, I)
'''''        matTImprimir(25, I) = matTImprimir(26, I) + matTImprimir(27, I)
'''''    Next
'''''
'''''    Open psArchivoAGrabar For Output As #1
'''''
'''''    Print #1, "01030100" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012"
'''''        sCad = ""
'''''
'''''        For I = 1 To 35
'''''            sCad = ""
'''''            For J = 2 To 20
'''''                If J = 2 Then
'''''                    If matTImprimir(I, J) > 0 Then
'''''                        sCad = sCad & Trim(Mid(Trim(Str(matTImprimir(I, J))), 2, 4))
'''''                    Else
'''''                        sCad = sCad & "0000"
'''''                    End If
'''''                Else
'''''                    If J = 3 Or J = 12 Then
'''''                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, J), 2)
'''''                    Else
'''''                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, J))
'''''                    End If
'''''                End If
'''''            Next
'''''
'''''            sCad = IIf(I = 1, " 100", IIf(I = 2, " 150", IIf(I > 2 And I < 11, " " & Trim(Str((I - 1) * 100)), Trim(Str((I - 1) * 100))))) & sCad
'''''
'''''            Print #1, sCad
'''''        Next
'''''
'''''
'''''
'''''    Close #1
'''''
'''''    ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, False
'''''
'''''    MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'''''
'''''    Exit Sub
'''''
'''''
'''''
'''''ErrBegin:
'''''  ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, True
'''''
'''''  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
'''''End Sub


'Private Sub GeneraSUCAVEAnx03(pnBitCentral As Boolean, pdFecha As Date, psMes As String)
'Dim psArchivoA_Leer As String
'Dim psArchivoAGrabar As String
'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja As Excel.Worksheet
'Dim bExiste As Boolean
'Dim bEncontrado As Boolean
'Dim fs As New scripting.FileSystemObject
'Dim matTImprimir(35, 20) As Currency
'Dim sCad As String
'Dim I As Integer
'Dim J As Integer
'Dim nFil As Integer
'Dim nCol As Integer
'
'On Error GoTo ErrBegin
'
''psArchivoA_Leer = App.path & "\Spooler\ANX3_" & Format(gdFecSis & " " & Time, "yyyymmdd") & gsCodUser & ".xls"
'
'psArchivoA_Leer = App.path & "\Spooler\Anx03_" & Mid(pdFecha, 7, 4) & Mid(pdFecha, 4, 2) & gsCodUser & ".xls"
'
'bExiste = fs.FileExists(psArchivoA_Leer)
'
'If bExiste = False Then
'    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoA_Leer, vbExclamation, "Aviso!!!"
'    Exit Sub
'End If
'
'    'Anexo 01 del 15A
'    '================
'
'    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".103"
'
'    Set xlAplicacion = New Excel.Application
'
'    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoA_Leer)
'    '''''''''''''''''''''''''''''
'    bEncontrado = False
'    For Each xlHoja In xlLibro.Worksheets
'        If UCase(xlHoja.Name) = UCase(psMes) Then
'            bEncontrado = True
'            xlHoja.Activate
'            Exit For
'        End If
'    Next
'
'    If bEncontrado = False Then
'        ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja, True
'        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
'        Exit Sub
'    End If
'    '''''''''''''''''''''''''''''
'    'Set xlHoja = xlAplicacion.Worksheets(1)
'
'
'
'    For I = 1 To 35
'        If I = 1 Then
'            nFil = I + 11
'        ElseIf I >= 2 And I <= 4 Then
'            nFil = I + 11
'        ElseIf I = 5 Then 'Aca se deben sumar desede Alimentos, Bebidas y Tabaco hasta RestoManufactura
'        ElseIf I >= 6 And I <= 15 Then
'            nFil = I + 10
'        ElseIf I >= 16 And I <= 17 Then
'            nFil = I + 10
'        ElseIf I = 18 Then 'Aca se deben sumar desde Venta y Reparacion de Vehiculos hasta comercio al por menor
'        ElseIf I >= 19 And I <= 21 Then
'            nFil = I + 9
'        ElseIf I >= 22 And I <= 24 Then
'            nFil = I + 9
'        ElseIf I = 25 Then 'Aca se deben sumar desde Actividad Inmoviliaria y de Alquiler hasta actividad empresarial
'        ElseIf I >= 26 And I <= 27 Then
'            nFil = I + 8
'        ElseIf I >= 28 And I <= 32 Then
'            nFil = I + 8
'        ElseIf I = 33 Then
'            nFil = 44
'        ElseIf I = 34 Then
'            nFil = 42
'        ElseIf I = 35 Then
'            nFil = 48
'        End If
'
'        If I <> 5 And I <> 18 And I <> 25 Then
'            For J = 2 To 20
'                If J = 2 Then
'                    If nFil <> 12 And nFil <> 42 And nFil <> 44 And nFil <> 48 Then
'                        nCol = J
'                        matTImprimir(I, J) = "1" & Mid(Trim(xlHoja.Cells(nFil, nCol)), 1, 2) & Mid(Trim(xlHoja.Cells(nFil, nCol)), 6, 2)
'                    End If
'                Else
'                    nCol = J
'                    matTImprimir(I, J) = xlHoja.Cells(nFil, nCol)
'                End If
'            Next
'        End If
'    Next
'
'    For I = 3 To 20
'        matTImprimir(5, I) = matTImprimir(6, I) + matTImprimir(7, I) + matTImprimir(8, I) + matTImprimir(9, I) + matTImprimir(10, I) + matTImprimir(11, I) + matTImprimir(12, I) + matTImprimir(13, I) + matTImprimir(14, I) + matTImprimir(15, I)
'        matTImprimir(18, I) = matTImprimir(19, I) + matTImprimir(20, I) + matTImprimir(21, I)
'        matTImprimir(25, I) = matTImprimir(26, I) + matTImprimir(27, I)
'    Next
'
'    Open psArchivoAGrabar For Output As #1
'
'    Print #1, "01030100" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012"
'        sCad = ""
'
'        For I = 1 To 35
'            sCad = ""
'            For J = 2 To 20
'                If J = 2 Then
'                    If matTImprimir(I, J) > 0 Then
'                        sCad = sCad & Trim(Mid(Trim(Str(matTImprimir(I, J))), 2, 4))
'                    Else
'                        sCad = sCad & "0000"
'                    End If
'                Else
'                    If J = 3 Or J = 12 Then
'                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, J), 2)
'                    Else
'                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, J))
'                    End If
'                End If
'            Next
'
'            sCad = IIf(I = 1, " 100", IIf(I = 2, " 150", IIf(I > 2 And I < 11, " " & Trim(Str((I - 1) * 100)), Trim(Str((I - 1) * 100))))) & sCad
'
'            Print #1, sCad
'        Next
'
'
'
'    Close #1
'
'    ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, False
'
'    MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'
'    Exit Sub
'
'
'
'ErrBegin:
'  ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, True
'
'  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
'End Sub



Private Sub GeneraSUCAVEAnx03(pnBitCentral As Boolean, pdFecha As Date, psMes As String)
Dim psArchivoA_Leer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim matTImprimir(35, 20) As Currency
Dim sCad As String
Dim I As Integer
Dim j As Integer
Dim nFil As Integer
Dim nCol As Integer

On Error GoTo ErrBegin

'psArchivoA_Leer = App.path & "\Spooler\ANX3_" & Format(gdFecSis & " " & Time, "yyyymmdd") & gsCodUser & ".xls"

psArchivoA_Leer = App.path & "\Spooler\Anx03_" & Mid(pdFecha, 7, 4) & Mid(pdFecha, 4, 2) & gsCodUser & ".xls"

bExiste = fs.FileExists(psArchivoA_Leer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoA_Leer, vbExclamation, "Aviso!!!"
    Exit Sub
End If

    'Anexo 01 del 15A
    '================

    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".103"

    Set xlAplicacion = New Excel.Application

    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoA_Leer)
    '''''''''''''''''''''''''''''
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase(psMes) Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    'Set xlHoja = xlAplicacion.Worksheets(1)



    For I = 1 To 35
        If I <= 32 Then
            nFil = I + 11
        ElseIf I = 33 Then
            nFil = I + 12
        ElseIf I = 34 Then
            nFil = I + 13
        ElseIf I = 35 Then
            nFil = I + 16
        End If
        
        
'        MsgBox "i=" & i & " Valor=" & xlHoja.Cells(nFil, 3)
        
'        If I = 1 Then
'            nFil = I + 11
'        ElseIf I >= 2 And I <= 4 Then
'            nFil = I + 11
'        ElseIf I = 5 Then 'Aca se deben sumar desede Alimentos, Bebidas y Tabaco hasta RestoManufactura
'        ElseIf I >= 6 And I <= 15 Then
'            nFil = I + 10
'        ElseIf I >= 16 And I <= 17 Then
'            nFil = I + 10
'        ElseIf I = 18 Then 'Aca se deben sumar desde Venta y Reparacion de Vehiculos hasta comercio al por menor
'        ElseIf I >= 19 And I <= 21 Then
'            nFil = I + 9
'        ElseIf I >= 22 And I <= 24 Then
'            nFil = I + 9
'        ElseIf I = 25 Then 'Aca se deben sumar desde Actividad Inmoviliaria y de Alquiler hasta actividad empresarial
'        ElseIf I >= 26 And I <= 27 Then
'            nFil = I + 8
'        ElseIf I >= 28 And I <= 32 Then
'            nFil = I + 8
'        ElseIf I = 33 Then
'            nFil = 44
'        ElseIf I = 34 Then
'            nFil = 42
'        ElseIf I = 35 Then
'            nFil = 48
'        End If
        
        If I <> 5 And I <> 18 And I <> 25 Then
            For j = 2 To 20
                If j = 2 Then
                    If nFil <> 12 And nFil <> 42 And nFil <> 44 And nFil <> 48 Then
                        nCol = j
                        matTImprimir(I, j) = "1" & Mid(Trim(xlHoja.Cells(nFil, nCol)), 1, 2) & Mid(Trim(xlHoja.Cells(nFil, nCol)), 6, 2)
                    End If
                Else
                    nCol = j
                    matTImprimir(I, j) = xlHoja.Cells(nFil, nCol)
                End If
            Next
        End If
    Next
    
    For I = 3 To 20
        matTImprimir(5, I) = matTImprimir(6, I) + matTImprimir(7, I) + matTImprimir(8, I) + matTImprimir(9, I) + matTImprimir(10, I) + matTImprimir(11, I) + matTImprimir(12, I) + matTImprimir(13, I) + matTImprimir(14, I) + matTImprimir(15, I)
        matTImprimir(18, I) = matTImprimir(19, I) + matTImprimir(20, I) + matTImprimir(21, I)
        matTImprimir(25, I) = matTImprimir(26, I) + matTImprimir(27, I)
    Next
  
    Open psArchivoAGrabar For Output As #1

    Print #1, "01030100" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012"
        sCad = ""
    
        For I = 1 To 35
            sCad = ""
            For j = 2 To 20
                If j = 2 Then
                    If matTImprimir(I, j) > 0 Then
                        'sCad = sCad & Trim(Mid(Trim(Str(matTImprimir(i, J))), 2, 4))
                        sCad = sCad & Left(Trim(Mid(Trim(Str(matTImprimir(I, j))), 2, 4)) & "     ", 4)
                    Else
                        sCad = sCad & "0000"
                    End If
                Else
                    If j = 3 Then
                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, j), 5)
                    Else
                        sCad = sCad & LlenaCerosSUCAVE(matTImprimir(I, j))
                    End If
                End If
            Next
            
            sCad = IIf(I = 1, " 100", IIf(I = 2, " 150", IIf(I > 2 And I < 11, " " & Trim(Str((I - 1) * 100)), Trim(Str((I - 1) * 100))))) & sCad
            
            Print #1, sCad
        Next
        

    
    Close #1

    ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, False

    MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
    Exit Sub



ErrBegin:
  ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub


Private Sub GeneraReporteAnexo5(pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, pnTipoMoneda As Integer)
Dim I As Integer
Dim k As Integer
Dim j As Integer
Dim nFila As Integer
Dim nIni  As Integer
Dim lNegativo As Boolean
Dim ldFecha   As Date
Dim sConec As String
Dim sTipoGara As String
Dim sTipoCred As String

    ldFecha = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
    CabeceraExcelAnexo5 ldFecha, psMes, pnTipoMoneda
  '  If Not oCon.AbreConexionRemota(Right(gsCodAge, 2), True, False, "03") Then
   '    Exit Sub
    'End If
    If oCon.AbreConexion = False Then Exit Sub

    oBarra.Max = 14
    oBarra.Progress 0, lsTitulo, "Generando reporte...", , vbBlue
    GetFileAnexos5 ldFecha, pnTipoMoneda
    oBarra.Progress 3, lsTitulo, "Generando reporte...", , vbBlue

    TotalesExcel5 11, 14
    TotalesExcel5 16, 20 '19
    TotalesExcel5 22, 25
    TotalesExcel5 27, 30
    TotalesExcel5 32, 34
    oBarra.Progress 5, lsTitulo, "Generando reporte...", , vbBlue
    TotalesExcel5 36, 38
    TotalesExcel5 40, 41
    TotalesExcel5 43, 44
    TotalesExcel5 46, 54
    TotalesExcel5 56, 59
    oBarra.Progress 7, lsTitulo, "Generando reporte...", , vbBlue
    TotalesExcel5 61, 64
    TotalesExcel5 66, 69
    TotalesExcel5 71, 74
    TotalesExcel5 80, 83
    TotalesExcel5 85, 88
    TotalesExcel5 90, 93
    TotalesExcel5 95, 98
    oBarra.Progress 9, lsTitulo, "Generando reporte...", , vbBlue
    TotalesExcel5 104, 107
    TotalesExcel5 109, 112
    TotalesExcel5 114, 117
    TotalesExcel5 119, 122
    
    oBarra.Progress 11, lsTitulo, "Generando reporte...", , vbBlue

    'TotalesExcel5C 20, 16, 19
    TotalesExcel5C 75, 71, 74
    TotalesExcel5C 99, 95, 98
    
    oBarra.Progress 13, lsTitulo, "Totalizando Moneda Extranjera...", , vbBlue

    xlHoja1.Range("E129:E129").Formula = "=+H11 + H12 + H13 + H14"
    xlHoja1.Range("F129:F129").Formula = "=+C61 + C62 + C63 + C64 + C85 + C86 + C87 + C88 + C109 + C110 + C111 + C112"
    xlHoja1.Range("G129:G129").Formula = "=+H61 + H62 + H63 + H64 + H85 + H86 + H87 + H88 + H109 + H110 + H111 + H112 - F129"

    FormatoReporteFinal5
    
    oBarra.Progress 14, lsTitulo, "Generando reporte...", , vbBlue

   oCon.CierraConexion
   RSClose R
End Sub


Private Sub CabeceraExcelAnexo5(pdFecha As Date, psMes As String, pnTipoMoneda As Integer)

xlHoja1.PageSetup.Orientation = xlPortrait
xlHoja1.PageSetup.CenterVertically = True
xlHoja1.PageSetup.Zoom = 50

xlHoja1.Cells(2, 2) = "ANEXO NRO. 05"
xlHoja1.Cells(3, 2) = " INFORME DE CLASIFICACION DE DEUDORES Y PROVISIONES "
xlHoja1.Cells(4, 2) = "EMPRESA : " & gsNomCmac
xlHoja1.Cells(4, 8) = "CODIGO : " & gsCodCMAC
xlHoja1.Cells(5, 2) = "Al " & Mid(pdFecha, 1, 2) & " de " & Trim(psMes) & " de " & Year(pdFecha)
If pnTipoMoneda = 1 Then
    xlHoja1.Cells(6, 2) = "( En Miles de Nuevos Soles )"
Else
    xlHoja1.Cells(6, 2) = "( En Nuevos Soles )"
End If
xlHoja1.Cells(8, 2) = "I.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA DE CREDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS 1/"
xlHoja1.Cells(77, 2) = "II.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA TRANSFERIDA SEGUN D.S.N° 114-98-EF 23/"
xlHoja1.Cells(101, 2) = "III.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA TRANSFERIDA SEGUN D.S.N° 099-99-EF   28/"
xlHoja1.Cells(125, 2) = "IV.- CUADRE DEL ANEXO N°5 CON CIFRAS DEL BALANCE  33/"

xlHoja1.Cells(127, 2) = "V.- CCIFRAS DE BALANCE"
xlHoja1.Cells(127, 5) = "W.- ANEXO N°5"

xlHoja1.Range("B2:H2").Merge
xlHoja1.Range("B3:H3").Merge
xlHoja1.Range("B5:H5").Merge
xlHoja1.Range("B6:H6").Merge
xlHoja1.Range("B8:H8").Merge
xlHoja1.Range("B77:H77").Merge
xlHoja1.Range("B101:H101").Merge
xlHoja1.Range("B125:H125").Merge
xlHoja1.Range("B127:D127").Merge
xlHoja1.Range("E127:G127").Merge

xlHoja1.Range("B2:H2").HorizontalAlignment = xlCenter
xlHoja1.Range("B3:H3").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:H5").HorizontalAlignment = xlCenter
xlHoja1.Range("B6:H6").HorizontalAlignment = xlCenter
xlHoja1.Range("B8:H8").HorizontalAlignment = xlCenter
xlHoja1.Range("B8:H8").VerticalAlignment = xlCenter
xlHoja1.Range("B8:H8").WrapText = True
xlHoja1.Range("B77:H77").HorizontalAlignment = xlCenter
xlHoja1.Range("B77:H77").VerticalAlignment = xlCenter
xlHoja1.Range("B77:H77").WrapText = True
xlHoja1.Range("B101:H101").HorizontalAlignment = xlCenter
xlHoja1.Range("B101:H101").VerticalAlignment = xlCenter
xlHoja1.Range("B101:H101").WrapText = True
xlHoja1.Range("B125:H125").HorizontalAlignment = xlCenter
xlHoja1.Range("B125:H125").VerticalAlignment = xlCenter
xlHoja1.Range("B125:H125").WrapText = True
xlHoja1.Range("B127:D127").HorizontalAlignment = xlCenter
xlHoja1.Range("B127:D127").VerticalAlignment = xlCenter
xlHoja1.Range("B127:D127").WrapText = True
xlHoja1.Range("E127:G127").HorizontalAlignment = xlCenter
xlHoja1.Range("E127:G127").VerticalAlignment = xlCenter
xlHoja1.Range("E127:G127").WrapText = True

If oCon.AbreConexion = False Then Exit Sub

GetFileAnexos5TIT

xlHoja1.Range("A1:A1").ColumnWidth = 1
xlHoja1.Range("B1:B1").ColumnWidth = 45
xlHoja1.Range("C1:C1").ColumnWidth = 20
xlHoja1.Range("D1:D1").ColumnWidth = 20
xlHoja1.Range("E1:E1").ColumnWidth = 20
xlHoja1.Range("F1:F1").ColumnWidth = 20
xlHoja1.Range("G1:G1").ColumnWidth = 20
xlHoja1.Range("H1:H1").ColumnWidth = 20

xlHoja1.Range("B128:G128").Font.Size = 8
xlHoja1.Range("B10:H123").Font.Size = 10.5
If pnTipoMoneda = 1 Then
    xlHoja1.Range("C11:H14").NumberFormat = "####,##0"
    xlHoja1.Range("C16:H20").NumberFormat = "####,##0"
    xlHoja1.Range("C22:H25").NumberFormat = "####,##0"
    xlHoja1.Range("C27:H30").NumberFormat = "####,##0"
    xlHoja1.Range("C32:H34").NumberFormat = "####,##0"
    xlHoja1.Range("C36:H38").NumberFormat = "####,##0"
    xlHoja1.Range("C40:H41").NumberFormat = "####,##0"
    xlHoja1.Range("C43:H44").NumberFormat = "####,##0"
    xlHoja1.Range("C46:H54").NumberFormat = "####,##0"
    xlHoja1.Range("C56:H59").NumberFormat = "####,##0"
    xlHoja1.Range("C61:H64").NumberFormat = "####,##0"
    xlHoja1.Range("C66:H69").NumberFormat = "####,##0"
    xlHoja1.Range("C71:H75").NumberFormat = "####,##0"
    xlHoja1.Range("C80:H83").NumberFormat = "####,##0"
    xlHoja1.Range("C85:H88").NumberFormat = "####,##0"
    xlHoja1.Range("C90:H93").NumberFormat = "####,##0"
    xlHoja1.Range("C95:H99").NumberFormat = "####,##0"
    xlHoja1.Range("C104:H107").NumberFormat = "####,##0"
    xlHoja1.Range("C109:H112").NumberFormat = "####,##0"
    xlHoja1.Range("C114:H117").NumberFormat = "####,##0"
    xlHoja1.Range("C119:H123").NumberFormat = "####,##0"
    xlHoja1.Range("B129:G129").NumberFormat = "####,##0"
Else
    xlHoja1.Range("C11:H14").NumberFormat = "####,##0.00"
    xlHoja1.Range("C16:H20").NumberFormat = "####,##0"
    xlHoja1.Range("C22:H25").NumberFormat = "####,##0.00"
    xlHoja1.Range("C27:H30").NumberFormat = "####,##0.00"
    xlHoja1.Range("C32:H34").NumberFormat = "####,##0.00"
    xlHoja1.Range("C36:H38").NumberFormat = "####,##0.00"
    xlHoja1.Range("C40:H41").NumberFormat = "####,##0.00"
    xlHoja1.Range("C43:H44").NumberFormat = "####,##0.00"
    xlHoja1.Range("C46:H54").NumberFormat = "####,##0.00"
    xlHoja1.Range("C56:H59").NumberFormat = "####,##0.00"
    xlHoja1.Range("C61:H64").NumberFormat = "####,##0.00"
    xlHoja1.Range("C66:H69").NumberFormat = "####,##0.00"
    xlHoja1.Range("C71:H75").NumberFormat = "####,##0.00"
    xlHoja1.Range("C80:H83").NumberFormat = "####,##0.00"
    xlHoja1.Range("C85:H88").NumberFormat = "####,##0.00"
    xlHoja1.Range("C90:H93").NumberFormat = "####,##0.00"
    xlHoja1.Range("C95:H99").NumberFormat = "####,##0.00"
    xlHoja1.Range("C104:H107").NumberFormat = "####,##0.00"
    xlHoja1.Range("C109:H112").NumberFormat = "####,##0.00"
    xlHoja1.Range("C114:H117").NumberFormat = "####,##0.00"
    xlHoja1.Range("C119:H123").NumberFormat = "####,##0.00"
    xlHoja1.Range("B129:G129").NumberFormat = "####,##0.00"
End If
End Sub

Private Sub GetFileAnexos5(pdFecha As Date, pnTipoMoneda As Integer)
Dim pSql As String
Dim prs As ADODB.Recordset
Dim lncol As Integer
Dim pnRow As Integer
Dim Monto As Double 'Long

pSql = "SELECT nANEXO, nAÑO, nMES, nLINEA, nCOLUMNA, nMONTO  " _
        & " FROM Anexos " _
        & "WHERE nANEXO = 5 AND nSUBAN = 1 AND nAÑO = " & Year(pdFecha) & " AND nMES = " & Month(pdFecha)

Set prs = oCon.CargaRecordSet(pSql, adLockReadOnly)
Do While Not prs.EOF
    pnRow = prs!nlinea
    lncol = prs!nColumna
    If pnRow = 16 Or pnRow = 17 Or pnRow = 18 Or pnRow = 19 Or pnRow = 20 Then
      Monto = Int(prs!nMonto)
    Else
        If pnTipoMoneda = 1 Then
            Monto = Int(Round(prs!nMonto / 1000, 0))
        Else
            Monto = Round(prs!nMonto, 2)
        End If

'      Monto = Int(Round(prs!nMonto, 2))
    End If
    If xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "" Then
          xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = "="
    End If
    
    xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula = xlHoja1.Range(xlHoja1.Cells(pnRow, lncol), xlHoja1.Cells(pnRow, lncol)).Formula & "+" & Monto
    prs.MoveNext
Loop
RSClose prs
End Sub


Private Sub GetFileAnexos5TIT()
Dim pSql As String
Dim prs As ADODB.Recordset
Dim lncol As Integer
Dim pnRow As Integer
Dim sCol As String
Dim nFila As Integer
nFila = 9
pSql = "SELECT nLINEA, nCOLUMNA, nANCHO, cTITULO, nNEGRITA, nCENTRADO   " _
        & " FROM AnexosTIT " _
        & "WHERE nANEXO = 5 AND nSUBAN = 1"

Set prs = oCon.CargaRecordSet(pSql, adLockReadOnly)
Do While Not prs.EOF
    pnRow = prs!nlinea
    lncol = prs!nColumna
    sCol = ExcelColumnaString(lncol)
    xlHoja1.Cells(pnRow + nFila, lncol) = prs!cTITULO
    If prs!nCENTRADO = 1 Then
       xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).HorizontalAlignment = xlCenter
    Else
       xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).HorizontalAlignment = xlLeft
    End If
    xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).VerticalAlignment = xlCenter
    xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).WrapText = True
    If prs!nNEGRITA = 1 Then
       xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).Font.Bold = True
    End If
    xlHoja1.Range(sCol & pnRow + nFila & ":" & sCol & pnRow + nFila).Font.Size = 8
prs.MoveNext
Loop
RSClose prs
End Sub


Private Sub TotalesExcel5(DESDE As Integer, HASTA As Integer)
Dim nLin As Integer
For nLin = DESDE To HASTA
    xlHoja1.Range("H" & nLin & ":H" & nLin).Formula = "=SUM(C" & nLin & ":G" & nLin & ")"
Next
End Sub

Private Sub TotalesExcel5C(nFila As Integer, DESDE As Integer, HASTA As Integer)
Dim nCol As Integer
Dim sCol As String
For nCol = 3 To 8
    sCol = ExcelColumnaString(nCol)
    xlHoja1.Range(sCol & nFila & ":" & sCol & nFila).Formula = "=SUM(" & sCol & DESDE & ": " & sCol & HASTA & ")"
Next
End Sub
Private Sub FormatoReporteFinal5()

xlHoja1.Range("B10:H75").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B10:H75").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B10:H10").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B15:H15").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B15:H15").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B20:H20").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B20:H20").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B21:H21").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B26:H26").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B26:H26").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B31:H31").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B31:H31").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B35:H35").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B35:H35").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B39:H39").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B39:H39").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B42:H42").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B42:H42").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B45:H45").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B45:H45").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B52:H52").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B55:H55").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B55:H55").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B60:H60").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B60:H60").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B65:H65").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B65:H65").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B70:H70").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B70:H70").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B75:H75").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B79:H99").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B79:H99").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B79:H79").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B84:H84").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B84:H84").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B89:H89").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B89:H89").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B94:H94").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B94:H94").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B99:H99").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B103:H123").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B103:H123").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B103:H103").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B108:H108").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B108:H108").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B113:H113").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B113:H113").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B118:H118").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B118:H118").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("B123:H123").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B128:G129").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
xlHoja1.Range("B128:G129").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B128:G128").Borders(xlEdgeBottom).LineStyle = xlContinuous

xlHoja1.Range("D140:E140").MergeCells = True
xlHoja1.Range("D141:E141").MergeCells = True
xlHoja1.Range("G140:H140").MergeCells = True
xlHoja1.Range("G141:H141").MergeCells = True

xlHoja1.Cells(140, 2) = "Gerente"
xlHoja1.Cells(141, 2) = "General"
xlHoja1.Cells(140, 4) = "Contador"
xlHoja1.Cells(141, 4) = "General"
xlHoja1.Cells(140, 7) = "Funcionario"
xlHoja1.Cells(141, 7) = "Responsable"

xlHoja1.Range("B140:B140").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("D140:E140").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("G140:H140").Borders(xlEdgeTop).LineStyle = xlContinuous

xlHoja1.Range("B140:H141").HorizontalAlignment = xlCenter
xlHoja1.Range("B140:H141").Font.Size = 10
xlHoja1.Range("B140:H141").Font.Name = "Arial"
 
End Sub
'''''
'''''Private Sub GeneraReporteAnexo3(ByVal pdFecha As Date, ByVal pnTipCambio As Double, psMes As String)   ' Flujo Crediticio por Tipo de Credito
'''''Dim I As Integer
'''''Dim nfila As Integer
'''''Dim nIni  As Integer
'''''Dim lNegativo As Boolean
'''''Dim sConec As String
'''''Dim lsSql As String
'''''Dim rsRang As New ADODB.Recordset
'''''Dim lsCodRangINI() As String * 2
'''''Dim lsCodRangFIN() As String * 2
'''''Dim lsDesRang() As String
'''''
'''''Dim nTempoFila(1 To 3) As Integer
'''''
'''''Dim lnRangos As Integer
'''''Dim reg9 As New ADODB.Recordset
'''''Dim regCredInd As New ADODB.Recordset
'''''Dim lnNroDeudores As Long
'''''Dim lnSaldoMesAntSol As Currency, lnSaldoMesAntDol As Currency
'''''Dim lnSaldoSol As Currency, lnSaldoDol As Currency
'''''Dim lnDesembNueSol As Currency, lnDesembNueDol As Currency
'''''Dim lnDesembRefSol As Currency, lnDesembRefDol As Currency
'''''Dim ldFechaMesAnt As Date
'''''Dim CIIUReg As String
'''''Dim lnTipCambMesAnt As Currency
'''''Dim J As Integer
'''''Dim lnProduc As Integer
''''''Dim nFil As Integer
'''''
'''''Dim matFinMes(2, 4) As Currency
'''''Dim regTemp As New ADODB.Recordset
'''''Dim oConLocal As DConecta
'''''Dim nFilTemp As Integer
'''''Dim nTFilTemp As Integer
'''''Dim nTotalTemp(8) As Currency
'''''Dim nTTotalTemp(8) As Currency
'''''Dim nTmp As Integer
'''''Dim nTemp As Integer
'''''
'''''
'''''Dim sservidorconsolidada As String
'''''Dim rCargaRuta As ADODB.Recordset
'''''
'''''   ldFechaMesAnt = DateAdd("d", pdFecha, -1 * Day(pdFecha))
'''''   Dim oTC As New nTipoCambio
'''''   lnTipCambMesAnt = oTC.EmiteTipoCambio(ldFechaMesAnt + 1, TCFijoMes)
'''''
'''''    CabeceraExcelAnexo3 pdFecha, psMes
'''''
'''''   If Not oCon.AbreConexionRemota(Right(gsCodAge, 2), True, False, "03") Then
'''''      Exit Sub
'''''   End If
'''''
'''''    'Saldos de Fin de Mes para Creditos Comerciales y PYMES(1), Agricolas(2), de Consumo (3), Hipotecarios(4)
'''''
'''''    lsSql = " SELECT 1 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]0[12]%' " & _
'''''            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "')" & _
'''''            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'''''            " Union All " & _
'''''            " SELECT 2 AS Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]02060[12]02%' " & _
'''''            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''            " Where a.cCtaContCod = c.cCtaContCod " & _
'''''            " and CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'''''            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'''''            " Union All " & _
'''''            " SELECT 3 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]03%' " & _
'''''            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'''''            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'''''            " Union All " & _
'''''            " SELECT 4 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]04%' " & _
'''''            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'''''            " Group by SUBSTRING(C.cCtaContCod, 3,1)"
'''''    Set oConLocal = New DConecta
'''''    oConLocal.AbreConexion
'''''
'''''    'Agregado
'''''    Set rCargaRuta = oConLocal.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
'''''    If Not rCargaRuta.BOF Then
'''''        sservidorconsolidada = rCargaRuta!nConsSisValor
'''''    End If
'''''    Set rCargaRuta = Nothing
'''''    'Fin agregado
'''''
'''''    Set regTemp = oConLocal.CargaRecordSet(lsSql)
'''''    Do While Not regTemp.EOF
'''''        matFinMes(regTemp!nMoneda, regTemp!Tipo) = regTemp!nSaldoMN
'''''        regTemp.MoveNext
'''''    Loop
'''''    regTemp.Close
'''''    Set regTemp = Nothing
'''''    oConLocal.CierraConexion
'''''
''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''
'''''    'If gbBitCentral = True Then
'''''        lsSql = " select nDesde, nHasta, cDescrip from anxriesgosrango where copecod='770030' order by ccodrango"
'''''        Set oConLocal = New DConecta
'''''        oConLocal.AbreConexion
'''''        Set rsRang = oConLocal.CargaRecordSet(lsSql)
'''''
'''''    'Else
'''''    '    lsSql = " SELECT cCodtab, cNomtab, nRanIniTab, nRanFinTab, cValor FROM TablaCod " & _
'''''    '            " WHERE cCodTab like 'A7%' AND LEN(cCodTab) > 2 ORDER BY cCodTab"
'''''    '
'''''    '    Set rsRang = oCon.CargaRecordSet(lsSql)
'''''    '
'''''    'End If
'''''
'''''    If Not (rsRang.BOF And rsRang.EOF) Then
'''''        rsRang.MoveLast
'''''        ReDim lsCodRangINI(rsRang.RecordCount)
'''''        ReDim lsCodRangFIN(rsRang.RecordCount)
'''''        ReDim lsDesRang(rsRang.RecordCount)
'''''
'''''        lnRangos = rsRang.RecordCount
'''''        rsRang.MoveFirst
'''''        I = 0
'''''        CIIUReg = "("
'''''        Do While Not rsRang.EOF
'''''
'''''            'If gbBitCentral = True Then
'''''                lsDesRang(I) = Trim(rsRang!cDescrip)
'''''                lsCodRangINI(I) = FillNum(Str(rsRang!nDesde), 2, "0")
'''''                lsCodRangFIN(I) = FillNum(Str(rsRang!nHasta), 2, "0")
'''''                If lsCodRangINI(I) = lsCodRangFIN(I) Then
'''''                    CIIUReg = CIIUReg & "'" & lsCodRangINI(I) & "',"
'''''                Else
'''''                    For J = lsCodRangINI(I) To lsCodRangFIN(I)
'''''                        CIIUReg = CIIUReg & "'" & FillNum(Str(J), 2, "0") & "',"
'''''                    Next
'''''                End If
'''''            'Else
'''''            '    lsDesRang(I) = Trim(rsRang!cNomtab)
'''''            '    lsCodRangINI(I) = FillNum(Str(rsRang!nRanIniTab), 2, "0")
'''''            '    lsCodRangFIN(I) = FillNum(Str(rsRang!nRanFinTab), 2, "0")
'''''            '    If lsCodRangINI(I) = lsCodRangFIN(I) Then
'''''            '        CIIUReg = CIIUReg & "'" & lsCodRangINI(I) & "',"
'''''            '    Else
'''''            '        For J = lsCodRangINI(I) To lsCodRangFIN(I)
'''''            '            CIIUReg = CIIUReg & "'" & FillNum(Str(J), 2, "0") & "',"
'''''            '        Next
'''''            '    End If
'''''            'End If
'''''
'''''            I = I + 1
'''''            rsRang.MoveNext
'''''        Loop
'''''        CIIUReg = Left(CIIUReg, Len(CIIUReg) - 1) & ")"
'''''    End If
'''''
'''''    For I = 0 To lnRangos - 1
'''''        If I = 0 Or (lsCodRangINI(I) = "52" And lsCodRangFIN(I) = "52") Or (lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "52") Then
'''''            xlHoja1.Cells(I + 13, 1) = lsDesRang(I)
'''''
'''''            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
'''''                xlHoja1.Cells(I + 13, 2) = "'" & Trim(lsCodRangINI(I))
'''''            Else
'''''                xlHoja1.Cells(I + 13, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
'''''            End If
'''''
'''''            If lsCodRangINI(I) = "52" And lsCodRangFIN(I) = "52" Then
'''''                nFilTemp = I + 13
'''''            ElseIf lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "52" Then
'''''                nTFilTemp = I + 13
'''''            End If
'''''
'''''        Else
'''''
'''''            '=================== SALDO ACTUAL =====================
'''''
''''''            lsSQL = " SELECT Count(c.cCodCta) as Numero, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' THEN nSaldoCap END ) AS SaldoCapSol, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' THEN nSaldoCap END ) AS SaldoCapDol, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' AND c.cRefinan = 'N'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' AND c.cRefinan = 'N'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' AND c.cRefinan = 'R'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''                  " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' AND c.cRefinan = 'R'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''                  " FROM CreditoConsol C LEFT JOIN  FuenteIngresoConsol FI " & _
'''''               " ON C.cNumFuente = FI.cNumFuente " & _
'''''               " WHERE  ( C.cEstado ='F' or (c.cEstado = 'V' And cCondCre = 'J') ) " & _
'''''               " AND c.nSaldoCap > 0 AND c.cCodcta like '__[12]%'  " & _
'''''               " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(i)) & "'" & _
'''''               " AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(i)) & "'"
'''''
'''''            If gbBitCentral = True Then
'''''                lsSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'N' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'N' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'R' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'R' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''                      " FROM CreditoSaldoConsol CSC INNER JOIN CreditoConsolTotal CCT ON CSC.cCtaCod = CCT.cCtaCod " & _
'''''                      " LEFT JOIN FuenteIngresoConsol FI ON CCT.cNumFuente = FI.cNumFuente " & _
'''''                      " WHERE  ( CSC.nPrdEstado IN(" & cVigente & ") or (CSC.nPrdEstado=" & gColocEstRecVigJud & ") ) " & _
'''''                      " AND CSC.nSaldoCap > 0 AND CSC.cCtaCod like '_____[12]%' AND SUBSTRING(CSC.cCtaCod,6,3) not in ('121','221') " & _
'''''                      " AND Substring(FI.cActEcon,2,2) >= '" & Trim(lsCodRangINI(I)) & "' AND Substring(FI.cActEcon,2,2) <= '" & Trim(lsCodRangFIN(I)) & "' " & _
'''''                      " AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')"
'''''            Else
'''''                lsSql = " SELECT Count(CSC.cCodCta) as Numero, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''                      " FROM CreditoSaldoConsol CSC INNER JOIN CreditoConsolTotal CCT ON CSC.cCodCta = CCT.cCodCta " & _
'''''                      " LEFT JOIN FuenteIngresoConsol FI ON CCT.cNumFuente = FI.cNumFuente " & _
'''''                      " WHERE  ( CSC.cEstado ='F' or (CSC.cEstado = 'V' And CCT.cCondCre = 'J') ) " & _
'''''                      " AND CSC.nSaldoCap > 0 AND CSC.cCodcta like '__[12]%' " & _
'''''                      " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(I)) & "' AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(I)) & "' " & _
'''''                      " AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')"
'''''            End If
'''''
'''''
'''''            Set reg9 = oCon.CargaRecordSet(lsSql)
'''''
'''''            lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'''''            lnSaldoSol = IIf(IsNull(reg9!SaldoCapSol), 0, reg9!SaldoCapSol)
'''''            lnSaldoDol = IIf(IsNull(reg9!SaldoCapDol), 0, reg9!SaldoCapDol * pnTipCambio)
'''''            lnDesembNueSol = IIf(IsNull(reg9!MontoDesembNueSol), 0, reg9!MontoDesembNueSol)
'''''            lnDesembNueDol = IIf(IsNull(reg9!MontoDesembNueDol), 0, reg9!MontoDesembNueDol * pnTipCambio)
'''''            lnDesembRefSol = IIf(IsNull(reg9!MontoDesembRefSol), 0, reg9!MontoDesembRefSol)
'''''            lnDesembRefDol = IIf(IsNull(reg9!MontoDesembRefDol), 0, reg9!MontoDesembRefDol * pnTipCambio)
'''''
'''''            reg9.Close
'''''
'''''            '***************************
'''''            'Numero de Deudores
'''''            '***************************
'''''            lsSql = "Select count(cPerscod) as Numero From ( "
'''''            lsSql = lsSql & "Select PP.cPersCod " & _
'''''            " FROM CreditoSaldoConsol CSC INNER JOIN CreditoConsolTotal CCT ON CSC.cCtaCod = CCT.cCtaCod " & _
'''''                    " Inner Join ProductoPersonaConsol PP ON CSC.cCtaCod = PP.cCtaCod AND  PP.nPrdPersRelac = 20 " & _
'''''                      " LEFT JOIN FuenteIngresoConsol FI ON CCT.cNumFuente = FI.cNumFuente " & _
'''''                      " WHERE  ( CSC.nPrdEstado IN(" & cVigente & ") or (CSC.nPrdEstado=" & gColocEstRecVigJud & ") ) " & _
'''''                      " AND CSC.nSaldoCap > 0 AND CSC.cCtaCod like '_____[12]%' AND SUBSTRING(CSC.cCtaCod,6,3) not in ('121','221') " & _
'''''                      " AND Substring(FI.cActEcon,2,2) >= '" & Trim(lsCodRangINI(I)) & "' AND Substring(FI.cActEcon,2,2) <= '" & Trim(lsCodRangFIN(I)) & "' " & _
'''''                      " AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')" & _
'''''                " Group by PP.cPerscod "
'''''            lsSql = lsSql & " ) as T "
'''''            Set reg9 = oCon.CargaRecordSet(lsSql)
'''''            If reg9.RecordCount > 0 Then
'''''                lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'''''            Else
'''''                lnNroDeudores = 0
'''''            End If
'''''            reg9.Close
'''''
'''''            '=================== SALDO MES ANTERIOR =====================
'''''
'''''            If gbBitCentral = True Then
'''''                lsSql = " SELECT Count(CS.cCtaCod) as Numero, " & _
'''''                      " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'''''                      " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'''''                      " FROM CreditoSaldoConsol CS JOIN ( SELECT cCtaCod, Max(cNumFuente) cNumFuente FROM CreditoConsolTotal ct GROUP BY cCtaCod ) c on CS.cCtaCod = C.cCtaCod " & _
'''''                      " LEFT JOIN  FuenteIngresoConsol FI  ON C.cNumFuente = FI.cNumFuente" & _
'''''                      " WHERE CONVERT(VARCHAR(8), CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'''''                      " AND CS.nPrdEstado IN(" & cVigente & ", " & gColocEstRecVigJud & ") " & _
'''''                      " AND CS.nSaldoCap > 0 AND CS.cCtaCod like '_____[12]%'  AND SUBSTRING(CS.cCtaCod,6,3) not in ('121','221')   " & _
'''''                      " AND Substring(FI.cActEcon,2,2) >= '" & Trim(lsCodRangINI(I)) & "'" & _
'''''                      " AND Substring(FI.cActEcon,2,2) <= '" & Trim(lsCodRangFIN(I)) & "'"
'''''            Else
'''''                lsSql = " SELECT Count(CS.cCodCta) as Numero, " & _
'''''                      " SUM( CASE WHEN substring(CS.cCodCta,6,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'''''                      " SUM( CASE WHEN substring(CS.cCodcta,6,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'''''                      " FROM CreditoSaldoConsol CS JOIN ( SELECT cCodCta, Max(cNumFuente) cNumFuente FROM CreditoConsolTotal ct GROUP BY cCodCta ) c on CS.ccodCta = C.cCodCta " & _
'''''                      " LEFT JOIN  FuenteIngresoConsol FI  ON C.cNumFuente = FI.cNumFuente" & _
'''''                      " WHERE CONVERT(VARCHAR(8), CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'''''                      " AND CS.cEstado in('F','V') " & _
'''''                      " AND CS.nSaldoCap > 0 AND CS.cCodcta like '__[12]%'   " & _
'''''                      " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(I)) & "'" & _
'''''                      " AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(I)) & "'"
'''''            End If
'''''
'''''            Set reg9 = oCon.CargaRecordSet(lsSql)
'''''            lnSaldoMesAntSol = IIf(IsNull(reg9!SaldoCapAntSol), 0, reg9!SaldoCapAntSol)
'''''            lnSaldoMesAntDol = IIf(IsNull(reg9!SaldoCapAntDol), 0, reg9!SaldoCapAntDol * lnTipCambMesAnt)
'''''
'''''            reg9.Close
'''''
'''''            xlHoja1.Cells(I + 13, 1) = lsDesRang(I)
'''''
'''''            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
'''''                xlHoja1.Cells(I + 13, 2) = "'" & Trim(lsCodRangINI(I))
'''''            Else
'''''                xlHoja1.Cells(I + 13, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
'''''            End If
'''''
'''''            xlHoja1.Cells(I + 13, 3) = lnNroDeudores
'''''            xlHoja1.Cells(I + 13, 4) = lnSaldoMesAntSol
'''''            xlHoja1.Cells(I + 13, 5) = lnSaldoMesAntDol
'''''            xlHoja1.Cells(I + 13, 6) = lnDesembNueSol + lnDesembRefSol
'''''            xlHoja1.Cells(I + 13, 7) = lnDesembNueDol + lnDesembRefDol
'''''
'''''            xlHoja1.Cells(I + 13, 8) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - lnSaldoSol
'''''            xlHoja1.Cells(I + 13, 9) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - lnSaldoDol
'''''
'''''            If lsCodRangINI(I) = "01" And lsCodRangFIN(I) = "02" Then 'Agricultura
'''''
'''''            ElseIf lsCodRangINI(I) = "52" And lsCodRangFIN(I) = "52" Then 'Comercio al por menor
'''''                nFilTemp = I + 13
'''''            ElseIf lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "52" Then 'Comercio al por menor
'''''                nTFilTemp = I + 13
'''''
'''''            Else    'Diferente a Agricultura
'''''                xlHoja1.Cells(I + 13, 10) = lnSaldoSol
'''''                xlHoja1.Cells(I + 13, 11) = lnSaldoDol
'''''
'''''                'If (Mid(lsDesRang(I), 1, 1) <> "-" And (lsCodRangINI(I) <> "50" And lsCodRangFIN(I) <> "52")) Or (lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "50") Or (lsCodRangINI(I) = "51" And lsCodRangFIN(I) = "51") Then
'''''                If (Mid(lsDesRang(I), 1, 1) <> "-") Or (lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "50") Or (lsCodRangINI(I) = "51" And lsCodRangFIN(I) = "51") Then
'''''
'''''                    nTotalTemp(0) = nTotalTemp(0) + lnNroDeudores
'''''                    nTotalTemp(1) = nTotalTemp(1) + lnSaldoMesAntSol
'''''                    nTotalTemp(2) = nTotalTemp(2) + lnSaldoMesAntDol
'''''                    nTotalTemp(3) = nTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'''''                    nTotalTemp(4) = nTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'''''                    nTotalTemp(5) = nTotalTemp(5) + lnSaldoSol
'''''                    nTotalTemp(6) = nTotalTemp(6) + lnSaldoDol
'''''                    nTotalTemp(7) = nTotalTemp(7) + lnDesembNueSol
'''''                    nTotalTemp(8) = nTotalTemp(8) + lnDesembNueDol
'''''                End If
'''''
'''''                If (lsCodRangINI(I) = "50" And lsCodRangFIN(I) = "50") Or (lsCodRangINI(I) = "51" And lsCodRangFIN(I) = "51") Then
'''''                    nTTotalTemp(0) = nTTotalTemp(0) + lnNroDeudores
'''''                    nTTotalTemp(1) = nTTotalTemp(1) + lnSaldoMesAntSol
'''''                    nTTotalTemp(2) = nTTotalTemp(2) + lnSaldoMesAntDol
'''''                    nTTotalTemp(3) = nTTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'''''                    nTTotalTemp(4) = nTTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'''''                    nTTotalTemp(5) = nTTotalTemp(5) + lnSaldoSol
'''''                    nTTotalTemp(6) = nTTotalTemp(6) + lnSaldoDol
'''''                    nTTotalTemp(7) = nTTotalTemp(7) + lnDesembNueSol
'''''                    nTTotalTemp(8) = nTTotalTemp(8) + lnDesembNueDol
'''''                End If
'''''
'''''
'''''
'''''            End If
'''''            '**************************
'''''            'Creditos Indirectos
'''''            '**************************
'''''            sSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nMontoApr END ) AS SaldoCapSol, " & _
'''''                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nMontoApr END ) AS SaldoCapDol " & _
'''''                      " FROM " & sservidorconsolidada & "CartaFianzaConsol CSC " & _
'''''                      " LEFT JOIN " & sservidorconsolidada & "FuenteIngresoConsol FI ON CSC.cNumFuente = FI.cNumFuente " & _
'''''                      " WHERE  ( CSC.nPrdEstado IN(" & cVigente & ") or (CSC.nPrdEstado=" & gColocEstRecVigJud & ") ) " & _
'''''                      " AND CSC.nMontoApr > 0 AND CSC.cCtaCod like '_____[12]%' " & _
'''''                      " AND Substring(FI.cActEcon,2,2) >= '" & Trim(lsCodRangINI(I)) & "' AND Substring(FI.cActEcon,2,2) <= '" & Trim(lsCodRangFIN(I)) & "' "
'''''
'''''            Set regCredInd = New ADODB.Recordset
'''''            Set regCredInd = oCon.CargaRecordSet(sSql)
'''''
'''''            xlHoja1.Cells(I + 13, 12) = IIf(IsNull(regCredInd!Numero), 0, regCredInd!Numero)
'''''            xlHoja1.Cells(I + 13, 13) = IIf(IsNull(regCredInd!SaldoCapSol), 0, regCredInd!SaldoCapSol)
'''''            xlHoja1.Cells(I + 13, 14) = IIf(IsNull(regCredInd!SaldoCapDol), 0, regCredInd!SaldoCapDol * pnTipCambio)
'''''
'''''            regCredInd.Close
'''''            Set regCredInd = Nothing
'''''
'''''            '***************************
'''''            'Numero de Deudores
'''''            '***************************
'''''            lsSql = " Select Count(cPersCod) as Numero From ( "
'''''            lsSql = lsSql & " Select PP.cPersCod  " & _
'''''            " FROM " & sservidorconsolidada & "CartaFianzaConsol CSC " & _
'''''            " Inner Join ProductoPersonaConsol PP ON CSC.cCtaCod = PP.cCtaCod AND  PP.nPrdPersRelac = 20 " & _
'''''                      " LEFT JOIN " & sservidorconsolidada & "FuenteIngresoConsol FI ON CSC.cNumFuente = FI.cNumFuente " & _
'''''                      " WHERE  ( CSC.nPrdEstado IN(" & cVigente & ") or (CSC.nPrdEstado=" & gColocEstRecVigJud & ") ) " & _
'''''                      " AND CSC.nMontoApr > 0 AND CSC.cCtaCod like '_____[12]%' " & _
'''''                      " AND Substring(FI.cActEcon,2,2) >= '" & Trim(lsCodRangINI(I)) & "' AND Substring(FI.cActEcon,2,2) <= '" & Trim(lsCodRangFIN(I)) & "' " & _
'''''            " Group by PP.cPerscod "
'''''            lsSql = lsSql & " ) as T  "
'''''
'''''            Set reg9 = oCon.CargaRecordSet(lsSql)
'''''
'''''            If reg9.RecordCount > 0 Then
'''''                xlHoja1.Cells(I + 13, 12) = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'''''            Else
'''''                xlHoja1.Cells(I + 13, 12) = "0"
'''''            End If
'''''            reg9.Close
'''''
'''''
'''''            xlHoja1.Cells(I + 13, 15) = lnDesembNueSol
'''''            xlHoja1.Cells(I + 13, 16) = lnDesembNueDol
'''''            xlHoja1.Cells(I + 13, 17) = 0
'''''            xlHoja1.Cells(I + 13, 18) = 0
'''''            xlHoja1.Cells(I + 13, 19) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - lnSaldoSol
'''''            xlHoja1.Cells(I + 13, 20) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - lnSaldoDol
'''''
'''''        End If
'''''    Next I
'''''
'''''   I = I + 1
'''''
'''''Dim nOrden As Integer
'''''
'''''For lnProduc = 1 To 4
'''''    '=================== SALDO ACTUAL =====================
'''''
'''''    'lsSQL = " SELECT Count(c.cCodCta) as Numero, " & _
'''''          " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' THEN nSaldoCap END ) AS SaldoCapSol, " & _
'''''          " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' THEN nSaldoCap END ) AS SaldoCapDol, " & _
'''''          " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' AND c.cRefinan = 'N'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'''''          " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' AND c.cRefinan = 'N'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''          " SUM( CASE WHEN substring(c.cCodCta,6,1) = '1' AND c.cRefinan = 'R'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''          " SUM( CASE WHEN substring(c.cCodcta,6,1) = '2' AND c.cRefinan = 'R'  And (convert(varchar(8), dFecVig, 112) Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''          " FROM CreditoConsol C " & _
'''''          " WHERE  ( C.cEstado in('F','1','4','6','7') or (c.cEstado = 'V' And cCondCre = 'J') ) " & _
'''''          " AND c.nSaldoCap > 0 "
'''''
'''''        If gbBitCentral = True Then
'''''            lsSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CASE WHEN SUBSTRING(CSC.cCtaCod,6,3)='305' THEN 0 ELSE  nMontoDesemb END ) END ) "
'''''            If lnProduc = 4 Then
'''''                lsSql = lsSql & " + (Select Sum(c.nMonto)"
'''''                lsSql = lsSql & " From dbcmacica.dbo.Mov a"
'''''                lsSql = lsSql & " Join dbcmacica.dbo.MovCol b On a.nmovnro = b.nmovnro"
'''''                lsSql = lsSql & " Join dbcmacica.dbo.MovColDet c On c.nmovnro = b.nmovnro"
'''''                lsSql = lsSql & " And c.cOpeCod = b.cOpeCod And c.cCtaCod = b.cCtaCod And c.nNroCalen = b.nNroCalen"
'''''                lsSql = lsSql & " Where a.cOpeCod in ('150201', '150202') And Substring(a.cMovNro,1,6) = '" & Trim(Str(Year(pdFecha))) & Right("00" & Trim(Str(Month(pdFecha))), 2) & "' "
'''''                lsSql = lsSql & " And c.nPrdConceptoCod = 8000 And a.nMovFlag = 0 ) "
'''''            End If
'''''                lsSql = lsSql & "  AS MontoDesembNueSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''                    " FROM CreditoSaldoConsol CSC INNER JOIN CreditoConsolTotal CCT ON CSC.cCtaCod = CCT.cCtaCod " & _
'''''                    " WHERE  ( CSC.nPrdEstado in(" & cVigente & ", " & cPigno & ") or (CSC.nPrdEstado =" & gColocEstRecVigJud & ") ) " & _
'''''                    " AND (CSC.nSaldoCap > 0) AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "') "
'''''        Else
'''''            lsSql = " SELECT Count(CSC.cCodCta) as Numero, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'''''                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefDol " & _
'''''                    " FROM CreditoSaldoConsol CSC INNER JOIN CreditoConsolTotal CCT ON CSC.cCodCta = CCT.cCodCta " & _
'''''                    " WHERE  ( CSC.cEstado in('F','1','4','6','7') or (CSC.cEstado = 'V' And CCT.cCondCre = 'J') ) " & _
'''''                    " AND (CSC.nSaldoCap > 0) AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "') "
'''''        End If
'''''
''''''        If gbBitCentral = True Then
''''''            If lnProduc = 1 Then 'Agricultura
''''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____202%' "
''''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
''''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____[12]%' "
''''''            ElseIf lnProduc = 3 Then  ' Consumo
''''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____3%' "
''''''            ElseIf lnProduc = 4 Then  ' Hipotecarios
''''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____4%' "
''''''            End If
''''''        Else
''''''            If lnProduc = 1 Then 'Agricultura
''''''                lsSql = lsSql & " AND  CSC.cCodcta like '__202%' "
''''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
''''''                lsSql = lsSql & " AND  CSC.cCodcta like '__[12]%' "
''''''            ElseIf lnProduc = 3 Then  ' Consumo
''''''                lsSql = lsSql & " AND  CSC.cCodcta like '__3%' "
''''''            ElseIf lnProduc = 4 Then  ' Hipotecarios
''''''                lsSql = lsSql & " AND  CSC.cCodcta like '__4%' "
''''''            End If
''''''        End If
'''''
'''''        If gbBitCentral = True Then
'''''            If lnProduc = 1 Then 'Agricultura
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____202%' "
'''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____[12]%'  and SUBSTRING(CSC.cCtaCod,6,3) not in ('121','221') "
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____3%'  and SUBSTRING(CSC.cCtaCod,6,3) not in ('321') "
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____4%' "
'''''            End If
'''''        Else
'''''            If lnProduc = 1 Then 'Agricultura
'''''                lsSql = lsSql & " AND  CSC.cCodcta like '__202%' "
'''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = lsSql & " AND  CSC.cCodcta like '__[12]%' "
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                lsSql = lsSql & " AND  CSC.cCodcta like '__3%' "
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                lsSql = lsSql & " AND  CSC.cCodcta like '__4%' "
'''''            End If
'''''        End If
'''''
'''''    Set reg9 = oCon.CargaRecordSet(lsSql)
'''''
'''''    lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'''''    lnSaldoSol = IIf(IsNull(reg9!SaldoCapSol), 0, reg9!SaldoCapSol)
'''''    lnSaldoDol = IIf(IsNull(reg9!SaldoCapDol), 0, reg9!SaldoCapDol * pnTipCambio)
'''''    lnDesembNueSol = IIf(IsNull(reg9!MontoDesembNueSol), 0, reg9!MontoDesembNueSol)
'''''    lnDesembNueDol = IIf(IsNull(reg9!MontoDesembNueDol), 0, reg9!MontoDesembNueDol * pnTipCambio)
'''''    lnDesembRefSol = IIf(IsNull(reg9!MontoDesembRefSol), 0, reg9!MontoDesembRefSol)
'''''    lnDesembRefDol = IIf(IsNull(reg9!MontoDesembRefDol), 0, reg9!MontoDesembRefDol * pnTipCambio)
'''''
'''''    reg9.Close
'''''
'''''    '***************************
'''''    'Numero de Deudores
'''''    '***************************
'''''            lsSql = ""
'''''            If lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = " select SUM(nMONTO) Numero "
'''''                lsSql = lsSql & " from Anexos  where nANEXO = 5 and nAÑO = " & Str(Year(pdFecha)) & " and nMES = " & Right("00" & Trim(Str(Month(pdFecha))), 2)
'''''                lsSql = lsSql & " and nCOLUMNA >= 3 and  nCOLUMNA <= 7 and nLINEA in (16,17)"
'''''
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                'lsSql = " Select ( select SUM(nMONTO) "
'''''                'lsSql = lsSql & " from DBCmactAux..Anexos  where nANEXO = 5 and nAÑO = " & Str(Year(pdFecha)) & " and nMES = " & Right("00" & Trim(Str(Month(pdFecha))), 2)
'''''                'lsSql = lsSql & " and nCOLUMNA >= 3 and  nCOLUMNA <= 7 and nLINEA in (20) ) - "
'''''                'lsSql = lsSql & " (select SUM(nMONTO) "
'''''                'lsSql = lsSql & " from DBCmactAux..Anexos  where nANEXO = 5 and nAÑO = " & Str(Year(pdFecha)) & " and nMES = " & Right("00" & Trim(Str(Month(pdFecha))), 2)
'''''                'lsSql = lsSql & " and nCOLUMNA >= 3 and  nCOLUMNA <= 7 and nLINEA in (16,17,18) ) as Numero "
'''''
'''''                lsSql = " select SUM(nMONTO) Numero "
'''''                lsSql = lsSql & " from Anexos  where nANEXO = 5 and nAÑO = " & Str(Year(pdFecha)) & " and nMES = " & Right("00" & Trim(Str(Month(pdFecha))), 2)
'''''                lsSql = lsSql & " and nCOLUMNA >= 3 and  nCOLUMNA <= 7 and nLINEA in (19)"
'''''
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                lsSql = " select SUM(nMONTO) Numero "
'''''                lsSql = lsSql & " from Anexos  where nANEXO = 5 and nAÑO = " & Str(Year(pdFecha)) & " and nMES = " & Right("00" & Trim(Str(Month(pdFecha))), 2)
'''''                lsSql = lsSql & " and nCOLUMNA >= 3 and  nCOLUMNA <= 7 and nLINEA in (18)"
'''''
'''''            End If
'''''
'''''    If lnProduc > 1 Then
'''''        Set oConLocal = New DConecta
'''''        oConLocal.AbreConexion
'''''        Set reg9 = oConLocal.CargaRecordSet(lsSql)
'''''        If reg9.RecordCount > 0 Then
'''''            lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'''''        Else
'''''            lnNroDeudores = 0
'''''        End If
'''''        reg9.Close
'''''        Set oConLocal = Nothing
'''''    Else
'''''        lnNroDeudores = 0
'''''    End If
'''''    '=================== SALDO MES ANTERIOR =====================
'''''
'''''        If gbBitCentral = True Then
'''''            If lnProduc = 1 Then 'Agricultura
'''''                lsSql = ""
'''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = " SELECT 1 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''                        " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]0[12]%' " & _
'''''                        " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''                        " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(ldFechaMesAnt, "YYYYMMdd") & "')" & _
'''''                        " Group By SUBSTRING(C.cCtaContCod, 3,1) "
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                   lsSql = " SELECT 3 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''                            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]03%' " & _
'''''                            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''                            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(ldFechaMesAnt, "YYYYMMdd") & "') " & _
'''''                            " Group By SUBSTRING(C.cCtaContCod, 3,1) "
'''''
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                   lsSql = " SELECT 4 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'''''                            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[12][1456]04%' " & _
'''''                            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'''''                            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(ldFechaMesAnt, "YYYYMMdd") & "') " & _
'''''                            " Group by SUBSTRING(C.cCtaContCod, 3,1)"
'''''            End If
'''''        End If
'''''
'''''    If lnProduc > 1 Then
'''''        lnSaldoMesAntSol = 0
'''''        lnSaldoMesAntDol = 0
'''''        Set oConLocal = New DConecta
'''''        oConLocal.AbreConexion
'''''        Set reg9 = oConLocal.CargaRecordSet(lsSql)
'''''        Do While Not reg9.EOF
'''''            If reg9!nMoneda = 1 Then
'''''                lnSaldoMesAntSol = IIf(IsNull(reg9!nSaldoMN), 0, reg9!nSaldoMN)
'''''            Else
'''''                lnSaldoMesAntDol = IIf(IsNull(reg9!nSaldoMN), 0, reg9!nSaldoMN)
'''''            End If
'''''
'''''            reg9.MoveNext
'''''        Loop
'''''        reg9.Close
'''''        Set oConLocal = Nothing
'''''    Else
'''''        lnSaldoMesAntSol = 0
'''''        lnSaldoMesAntDol = 0
'''''    End If
'''''
'''''
'''''    '''If bandera = 1 Then
'''''
'''''    If lnProduc = 1 Then
'''''        nTmp = 0
'''''        'If gbBitCentral = True Then
'''''            xlHoja1.Cells(nTmp + 13, 1) = "A. Agricultura, Ganaderia, Caza y Silvicultura"
'''''        'Else
'''''        '    xlHoja1.Cells(nTmp + 13, 1) = "Agricultura"
'''''        'End If
'''''    ElseIf lnProduc = 2 Then
'''''        nTmp = -1
'''''        xlHoja1.Cells(nTmp + 13, 1) = "1. CRÉDITOS COMERCIALES Y MICROEMPRESAS"
'''''        xlHoja1.Cells(nTmp + 13, 2) = "'01 a 99"
'''''        nTempoFila(1) = nTmp + 13
'''''        xlHoja1.Range("A" & nTmp + 13 & ":S" & nTmp + 13).Font.Bold = True
'''''    Else
'''''        nTmp = I
'''''        xlHoja1.Cells(nTmp + 13, 1) = IIf(lnProduc = 4, "3. CREDITOS DE CONSUMO", "2. CREDITOS HIPOTECARIOS PARA VIVIENDA")
'''''        xlHoja1.Range("A" & nTmp + 13 & ":S" & nTmp + 13).Font.Bold = True
'''''        If lnProduc = 4 Then
'''''            nTempoFila(2) = nTmp + 13
'''''        Else
'''''            nTempoFila(3) = nTmp + 13
'''''        End If
'''''    End If
'''''
'''''    xlHoja1.Cells(nTmp + 13, 3) = lnNroDeudores
'''''    xlHoja1.Cells(nTmp + 13, 4) = lnSaldoMesAntSol
'''''    xlHoja1.Cells(nTmp + 13, 5) = lnSaldoMesAntDol
'''''    xlHoja1.Cells(nTmp + 13, 6) = lnDesembNueSol + lnDesembRefSol
'''''    xlHoja1.Cells(nTmp + 13, 7) = lnDesembNueDol + lnDesembRefDol
'''''
'''''    If lnProduc = 1 Then 'Agricultura
'''''
'''''        xlHoja1.Cells(nTmp + 13, 8) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 2)
'''''        xlHoja1.Cells(nTmp + 13, 9) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 2)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 10) = matFinMes(1, 2)
'''''        xlHoja1.Cells(nTmp + 13, 11) = matFinMes(2, 2)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 19) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 2)
'''''        xlHoja1.Cells(nTmp + 13, 20) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 2)
'''''
'''''        nTotalTemp(0) = nTotalTemp(0) + lnNroDeudores
'''''        nTotalTemp(1) = nTotalTemp(1) + lnSaldoMesAntSol
'''''        nTotalTemp(2) = nTotalTemp(2) + lnSaldoMesAntDol
'''''        nTotalTemp(3) = nTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'''''        nTotalTemp(4) = nTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'''''        nTotalTemp(5) = nTotalTemp(5) + matFinMes(1, 2)
'''''        nTotalTemp(6) = nTotalTemp(6) + matFinMes(2, 2)
'''''        nTotalTemp(7) = nTotalTemp(7) + lnDesembNueSol
'''''        nTotalTemp(8) = nTotalTemp(8) + lnDesembNueDol
'''''
'''''    ElseIf lnProduc = 2 Then ' Comerciales y PYMES
'''''
'''''        xlHoja1.Cells(nTmp + 13, 8) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)
'''''        xlHoja1.Cells(nTmp + 13, 9) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)
'''''
'''''
'''''        xlHoja1.Cells(nTmp + 13, 10) = matFinMes(1, 1)
'''''        xlHoja1.Cells(nTmp + 13, 11) = matFinMes(2, 1)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 19) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)
'''''        xlHoja1.Cells(nTmp + 13, 20) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)
'''''
'''''        'LLeno el comercio al por menor
'''''        xlHoja1.Cells(nFilTemp, 3) = Format(lnNroDeudores - nTotalTemp(0), "0")
'''''        xlHoja1.Cells(nFilTemp, 4) = lnSaldoMesAntSol - nTotalTemp(1)
'''''        xlHoja1.Cells(nFilTemp, 5) = lnSaldoMesAntDol - nTotalTemp(2)
'''''        xlHoja1.Cells(nFilTemp, 6) = lnDesembNueSol + lnDesembRefSol - nTotalTemp(3)
'''''        xlHoja1.Cells(nFilTemp, 7) = lnDesembNueDol + lnDesembRefDol - nTotalTemp(4)
'''''        xlHoja1.Cells(nFilTemp, 8) = (lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)) - (nTotalTemp(1) + nTotalTemp(3) - nTotalTemp(5))
'''''        xlHoja1.Cells(nFilTemp, 9) = (lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)) - (nTotalTemp(2) + nTotalTemp(4) - nTotalTemp(6))
'''''
'''''        xlHoja1.Cells(nFilTemp, 10) = matFinMes(1, 1) - nTotalTemp(5)
'''''        xlHoja1.Cells(nFilTemp, 11) = matFinMes(2, 1) - nTotalTemp(6)
'''''
'''''        '*******************************************
'''''        'Creditos Indirectos
'''''        '*******************************************
'''''        'lsSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'''''        '            " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nMontoApr END ) AS SaldoCapSol, " & _
'''''        '            " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nMontoApr END ) AS SaldoCapDol " & _
'''''        '            " FROM " & sservidorconsolidada & "CartaFianzaConsol CSC " & _
'''''        '            " WHERE  ( CSC.nPrdEstado in(" & cVigente & ", " & cPigno & ") or (CSC.nPrdEstado =" & gColocEstRecVigJud & ") ) " & _
'''''        '            " AND (CSC.nMontoApr > 0) "
'''''
'''''        '    If lnProduc = 1 Then 'Agricultura
'''''        '        lsSql = lsSql & " AND  CSC.cCtaCod like '_____202%' "
'''''        '    ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''        '        lsSql = lsSql & " AND  CSC.cCtaCod like '_____[12]%' "
'''''        '    ElseIf lnProduc = 4 Then  ' Consumo
'''''        '        lsSql = lsSql & " AND  CSC.cCtaCod like '_____3%' "
'''''        '    ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''        '        lsSql = lsSql & " AND  CSC.cCtaCod like '_____4%' "
'''''        '    End If
'''''
'''''        'Set regCredInd = oCon.CargaRecordSet(lsSql)
'''''
'''''        'xlHoja1.Cells(nFilTemp, 12) = IIf(IsNull(regCredInd!Numero), 0, regCredInd!Numero)
'''''        'xlHoja1.Cells(nFilTemp, 13) = IIf(IsNull(regCredInd!SaldoCapSol), 0, regCredInd!SaldoCapSol)
'''''        'xlHoja1.Cells(nFilTemp, 14) = IIf(IsNull(regCredInd!SaldoCapDol), 0, regCredInd!SaldoCapDol * pnTipCambio)
'''''
'''''        'regCredInd.Close
'''''        'Set regCredInd = Nothing
'''''
'''''
'''''        xlHoja1.Cells(nFilTemp, 15) = lnDesembNueSol - nTotalTemp(7)
'''''        xlHoja1.Cells(nFilTemp, 16) = lnDesembNueDol - nTotalTemp(8)
'''''
'''''        xlHoja1.Cells(nFilTemp, 17) = "0.00"
'''''        xlHoja1.Cells(nFilTemp, 18) = "0.00"
'''''
'''''        xlHoja1.Cells(nFilTemp, 19) = (lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)) - (nTotalTemp(1) + nTotalTemp(3) - nTotalTemp(5))
'''''        xlHoja1.Cells(nFilTemp, 20) = (lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)) - (nTotalTemp(2) + nTotalTemp(4) - nTotalTemp(6))
'''''
'''''        'Lleno los totales del grupo que agrupa a comercio exterior
'''''        xlHoja1.Cells(nTFilTemp, 3) = Format(lnNroDeudores - nTotalTemp(0) + nTTotalTemp(0), "0")
'''''        xlHoja1.Cells(nTFilTemp, 4) = lnSaldoMesAntSol - nTotalTemp(1) + nTTotalTemp(1)
'''''        xlHoja1.Cells(nTFilTemp, 5) = lnSaldoMesAntDol - nTotalTemp(2) + nTTotalTemp(2)
'''''        xlHoja1.Cells(nTFilTemp, 6) = lnDesembNueSol + lnDesembRefSol - nTotalTemp(3) + nTTotalTemp(3)
'''''        xlHoja1.Cells(nTFilTemp, 7) = lnDesembNueDol + lnDesembRefDol - nTotalTemp(4) + nTTotalTemp(4)
'''''        xlHoja1.Cells(nTFilTemp, 8) = (lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)) - (nTotalTemp(1) + nTotalTemp(3) - nTotalTemp(5)) + (nTTotalTemp(1) + nTTotalTemp(3) - nTTotalTemp(5))
'''''        xlHoja1.Cells(nTFilTemp, 9) = (lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)) - (nTotalTemp(2) + nTotalTemp(4) - nTotalTemp(6)) + (nTTotalTemp(2) + nTTotalTemp(4) - nTTotalTemp(6))
'''''
'''''        xlHoja1.Cells(nTFilTemp, 10) = matFinMes(1, 1) - nTotalTemp(5) + nTTotalTemp(5)
'''''        xlHoja1.Cells(nTFilTemp, 11) = matFinMes(2, 1) - nTotalTemp(6) + nTTotalTemp(6)
'''''
'''''        xlHoja1.Cells(nTFilTemp, 12) = "0.00"
'''''        xlHoja1.Cells(nTFilTemp, 13) = "0.00"
'''''        xlHoja1.Cells(nTFilTemp, 14) = "0.00"
'''''
'''''        xlHoja1.Cells(nTFilTemp, 15) = lnDesembNueSol - nTotalTemp(7) + nTTotalTemp(7)
'''''        xlHoja1.Cells(nTFilTemp, 16) = lnDesembNueDol - nTotalTemp(8) + nTTotalTemp(8)
'''''
'''''        xlHoja1.Cells(nTFilTemp, 17) = "0.00"
'''''        xlHoja1.Cells(nTFilTemp, 18) = "0.00"
'''''
'''''        xlHoja1.Cells(nTFilTemp, 19) = (lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 1)) - (nTotalTemp(1) + nTotalTemp(3) - nTotalTemp(5)) + (nTTotalTemp(1) + nTTotalTemp(3) - nTTotalTemp(5))
'''''        xlHoja1.Cells(nTFilTemp, 20) = (lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 1)) - (nTotalTemp(2) + nTotalTemp(4) - nTotalTemp(6)) + (nTTotalTemp(2) + nTTotalTemp(4) - nTTotalTemp(6))
'''''
'''''
'''''    ElseIf lnProduc = 4 Then 'De Consumo
'''''
'''''        xlHoja1.Cells(nTmp + 13, 8) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 3)
'''''        xlHoja1.Cells(nTmp + 13, 9) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 3)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 10) = matFinMes(1, 3)
'''''        xlHoja1.Cells(nTmp + 13, 11) = matFinMes(2, 3)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 19) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 3)
'''''        xlHoja1.Cells(nTmp + 13, 20) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 3)
'''''    ElseIf lnProduc = 3 Then 'Hipotecarios
'''''
'''''
'''''        xlHoja1.Cells(nTmp + 13, 8) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 4)
'''''        xlHoja1.Cells(nTmp + 13, 9) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 4)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 10) = matFinMes(1, 4)
'''''        xlHoja1.Cells(nTmp + 13, 11) = matFinMes(2, 4)
'''''
'''''        xlHoja1.Cells(nTmp + 13, 19) = lnSaldoMesAntSol + lnDesembNueSol + lnDesembRefSol - matFinMes(1, 4)
'''''        xlHoja1.Cells(nTmp + 13, 20) = lnSaldoMesAntDol + lnDesembNueDol + lnDesembRefDol - matFinMes(2, 4)
'''''    End If
'''''
'''''
'''''    '*******************************************
'''''        'Creditos Indirectos
'''''        '*******************************************
'''''        lsSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nMontoApr END ) AS SaldoCapSol, " & _
'''''                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nMontoApr END ) AS SaldoCapDol " & _
'''''                    " FROM " & sservidorconsolidada & "CartaFianzaConsol CSC " & _
'''''                    " WHERE  ( CSC.nPrdEstado in(" & cVigente & ", " & cPigno & ") or (CSC.nPrdEstado =" & gColocEstRecVigJud & ") ) " & _
'''''                    " AND (CSC.nMontoApr > 0) "
'''''
'''''            If lnProduc = 1 Then 'Agricultura
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____202%' "
'''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____[12]%' "
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____3%' "
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____4%' "
'''''            End If
'''''
'''''        Set regCredInd = oCon.CargaRecordSet(lsSql)
'''''
'''''    xlHoja1.Cells(nTmp + 13, 12) = IIf(IsNull(regCredInd!Numero), 0, regCredInd!Numero)
'''''    xlHoja1.Cells(nTmp + 13, 13) = IIf(IsNull(regCredInd!SaldoCapSol), 0, regCredInd!SaldoCapSol)
'''''    xlHoja1.Cells(nTmp + 13, 14) = IIf(IsNull(regCredInd!SaldoCapDol), 0, regCredInd!SaldoCapDol * pnTipCambio)
'''''
'''''    regCredInd.Close
'''''
'''''
'''''    '************************************
'''''    ' Numero de Deudores
'''''    '************************************
'''''    lsSql = " Select COUNT(cPersCod) as Numero From ( "
'''''    lsSql = lsSql & " Select PP.cPersCod " & _
'''''    " FROM " & sservidorconsolidada & "CartaFianzaConsol CSC " & _
'''''    " Inner Join ProductoPersonaConsol PP ON CSC.cCtaCod = PP.cCtaCod AND  PP.nPrdPersRelac = 20 " & _
'''''                    " WHERE  ( CSC.nPrdEstado in(" & cVigente & ", " & cPigno & ") or (CSC.nPrdEstado =" & gColocEstRecVigJud & ") ) " & _
'''''                    " AND (CSC.nMontoApr > 0) "
'''''
'''''            If lnProduc = 1 Then 'Agricultura
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____202%' "
'''''            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____[12]%' "
'''''            ElseIf lnProduc = 4 Then  ' Consumo
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____3%' and SUBSTRING(CSC.cCtaCod,6,3) not in ('321') AND PP.cPerscod not in ( Select cPerscod From ProductoPersonaConsol Where nPrdPersrelac = 20 and substring(cctaCod,6,1) in ('1','2','4') ) "
'''''            ElseIf lnProduc = 3 Then  ' Hipotecarios
'''''                lsSql = lsSql & " AND  CSC.cCtaCod like '_____4%' and SUBSTRING(CSC.cCtaCod,6,3) not in ('321') AND PP.cPerscod not in ( Select cPerscod From ProductoPersonaConsol Where nPrdPersrelac = 20 and substring(cctaCod,6,1) in ('1','2') ) "
'''''            End If
'''''    lsSql = lsSql & " Group by PP.cPersCod "
'''''    lsSql = lsSql & " ) as T "
'''''
'''''    Set regCredInd = oCon.CargaRecordSet(lsSql)
'''''    If regCredInd.RecordCount > 0 Then
'''''        xlHoja1.Cells(nTmp + 13, 12) = IIf(IsNull(regCredInd!Numero), 0, regCredInd!Numero)
'''''    Else
'''''        xlHoja1.Cells(nTmp + 13, 12) = "0"
'''''    End If
'''''    regCredInd.Close
'''''    Set regCredInd = Nothing
'''''
'''''
'''''    xlHoja1.Cells(nTmp + 13, 15) = lnDesembNueSol
'''''    xlHoja1.Cells(nTmp + 13, 16) = lnDesembNueDol
'''''
'''''    xlHoja1.Cells(nTmp + 13, 17) = 0
'''''    xlHoja1.Cells(nTmp + 13, 18) = 0
'''''
'''''
'''''    If lnProduc <> 1 Then
'''''        xlHoja1.Range("A" & nTmp + 13 & ":T" & nTmp + 13).Font.Bold = True
'''''    End If
'''''
'''''    If lnProduc = 1 Or lnProduc = 2 Then
'''''    Else
'''''        I = I + 2
'''''    End If
'''''
''''' Next
'''''
'''''    xlHoja1.Cells(I + 15, 1) = "Total"
'''''
'''''    ExcelCuadro xlHoja1, 1, I + 15, 20, I + 15
'''''
'''''    'xlHoja1.Range("C12:C" & I + 15).NumberFormat = "#,###,##0"
'''''
'''''    For nTemp = 99 To 116
'''''        xlHoja1.Range(UCase(Chr(nTemp)) & I + 15 & ":" & UCase(Chr(nTemp)) & I + 15).Formula = "=+" & UCase(Chr(nTemp)) & nTempoFila(1) & "+" & UCase(Chr(nTemp)) & nTempoFila(2) & "+" & UCase(Chr(nTemp)) & nTempoFila(3)
'''''    Next
'''''
'''''    xlHoja1.Range("A" & I + 15 & ":T" & I + 15).Font.Bold = True
'''''
'''''    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(I + 15, 20)).Borders(xlInsideVertical).LineStyle = xlContinuous
'''''    xlHoja1.Range("A7:T" & 15 + I).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
'''''
'''''    xlHoja1.Cells(I + 16, 1) = "Periodicidad Mensual"
'''''    xlHoja1.Cells(I + 18, 1) = "(1) Clasificación industrial uniforme de todas las Actividades económicas. Tercera Revisión. Naciones Unidas. Consignar la actividad económica que genera el mayor valor añadido de la entidad deudora"
'''''    xlHoja1.Cells(I + 19, 1) = "(2) El total de créditos directos debe coincidir con la suma de las cuentas 1401+1403+1404+1405+1406+1407 del Manual de Contabilidad"
'''''    xlHoja1.Cells(I + 20, 1) = "(3) El total de créditos indirectos debe corresponder a la suma de los saldos de las cuentas 7101+7102+7103+7104+7105 del Manual de Contabilidad"
'''''
'''''    xlHoja1.Range("A" & I + 16 & ":A" & I + 16).Font.Bold = True
'''''    xlHoja1.Range("A" & I + 16 & ":A" & I + 20).Font.Size = 8
'''''
'''''    xlHoja1.Range("D" & I + 22 & ":E" & I + 22).MergeCells = True
'''''    xlHoja1.Range("I" & I + 22 & ":J" & I + 22).MergeCells = True
'''''    xlHoja1.Range("O" & I + 22 & ":P" & I + 22).MergeCells = True
'''''
'''''    xlHoja1.Range("D" & I + 23 & ":E" & I + 23).MergeCells = True
'''''    xlHoja1.Range("I" & I + 23 & ":J" & I + 23).MergeCells = True
'''''    xlHoja1.Range("O" & I + 23 & ":P" & I + 23).MergeCells = True
'''''
'''''    xlHoja1.Range("I" & I + 24 & ":J" & I + 24).MergeCells = True
'''''
'''''    xlHoja1.Cells(I + 23, 4) = "Gerente General"
'''''    xlHoja1.Cells(I + 23, 9) = "Contador General"
'''''    xlHoja1.Cells(I + 24, 9) = "Matricula Nro"
'''''    xlHoja1.Cells(I + 23, 15) = "Hecho Por"
'''''
'''''    xlHoja1.Range("A" & I + 23 & ":T" & I + 24).Font.Bold = True
'''''    xlHoja1.Range("A" & I + 23 & ":T" & I + 24).Font.Size = 8
'''''
'''''
'''''    xlHoja1.Range("D" & I + 23 & ":P" & I + 24).HorizontalAlignment = xlCenter
'''''    xlHoja1.Range("D" & I + 23 & ":P" & I + 24).HorizontalAlignment = xlCenter
'''''
'''''    xlHoja1.Range("D" & I + 23 & ":E" & I + 23).Borders(xlEdgeTop).LineStyle = xlContinuous
'''''    xlHoja1.Range("I" & I + 23 & ":J" & I + 23).Borders(xlEdgeTop).LineStyle = xlContinuous
'''''    xlHoja1.Range("O" & I + 23 & ":P" & I + 23).Borders(xlEdgeTop).LineStyle = xlContinuous
'''''
'''''
'''''    xlHoja1.Range("C11:C" & I + 15).NumberFormat = "#,###,##0"
'''''    xlHoja1.Range("L11:L" & I + 15).NumberFormat = "#,###,##0"
'''''
'''''    If gbBitCentral = True Then
'''''
'''''        'oConLocal.CierraConexion
'''''    Else
'''''        oCon.CierraConexion
'''''    End If
'''''   frmMdiMain.staMain.Panels(1).Text = ""
'''''   RSClose R
'''''End Sub


'Private Sub GeneraReporteAnexo3(ByVal pdFecha As Date, ByVal pnTipCambio As Double, psMes As String)   ' Flujo Crediticio por Tipo de Credito
'Dim i As Integer
'Dim nFila As Integer
'Dim nIni  As Integer
'Dim lNegativo As Boolean
'Dim sConec As String
'Dim lsSql As String
'Dim rsRang As New ADODB.Recordset
'Dim lsCodRangINI() As String * 2
'Dim lsCodRangFIN() As String * 2
'Dim lsDesRang() As String
'
'Dim nTempoFila(1 To 3) As Integer
'
'Dim lnRangos As Integer
'Dim reg9 As New ADODB.Recordset
'Dim lnNroDeudores As Long
'Dim lnSaldoMesAntSol As Currency, lnSaldoMesAntDol As Currency
'Dim lnSaldoSol As Currency, lnSaldoDol As Currency
'Dim lnNumeroDesembNue As Long ' BRGO BASILEA II
'Dim lnDesembNueSol As Currency, lnDesembNueDol As Currency
'Dim lnDesembRefSol As Currency, lnDesembRefDol As Currency
'Dim ldFechaMesAnt As Date
'Dim CIIUReg As String
'Dim lnTipCambMesAnt As Currency
'Dim j As Integer
'Dim lnProduc As Integer
''Dim nFil As Integer
'
'Dim matFinMes(2, 4) As Currency
'Dim regTemp As New ADODB.Recordset
'Dim oConLocal As DConecta
'Dim nFilTemp As Integer
'Dim nTFilTemp As Integer
'Dim nTotalTemp(9) As Currency
'Dim nTTotalTemp(9) As Currency
'Dim nTmp As Integer
'Dim nTemp As Integer
'
'   ldFechaMesAnt = DateAdd("d", pdFecha, -1 * Day(pdFecha))
'   Dim oTC As New nTipoCambio
'   lnTipCambMesAnt = oTC.EmiteTipoCambio(ldFechaMesAnt + 1, TCFijoMes)
'
'   CabeceraExcelAnexo3 pdFecha, psMes
'
'   If Not oCon.AbreConexion Then 'Remota(Right(gsCodAge, 2), True, False, "03")
'      Exit Sub
'   End If
'
'    'Saldos de Fin de Mes para Creditos Corp,Gran,Med,Peq y Micro empresa (1), Agricolas(2), de Consumo (3), Hipotecarios(4)
'
'    lsSql = " SELECT 1 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'            " FROM CTASALDO C WHERE (C.cCtaContCod LIKE '14[132][1456]0[1256789]%' or C.cCtaContCod LIKE '14[132][1456]1[0123]%') " & _
'            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "')" & _
'            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'            " Union All " & _
'            " SELECT 2 AS Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'            " FROM CTASALDO C WHERE (C.cCtaContCod LIKE '14[132][1456]02060[129]02%' or C.cCtaContCod LIKE '14[132][1456]13060[12]02%') " & _
'            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'            " Where a.cCtaContCod = c.cCtaContCod " & _
'            " and CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'            " Union All " & _
'            " SELECT 3 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[132][1456]03%' " & _
'            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'            " Group By SUBSTRING(C.cCtaContCod, 3,1) " & _
'            " Union All " & _
'            " SELECT 4 as Tipo, SUBSTRING(C.cCtaContCod, 3,1) as nMoneda, SUM(C.nCtaSaldoImporte) AS nSaldoMN " & _
'            " FROM CTASALDO C WHERE C.cCtaContCod LIKE '14[132][1456]04%' " & _
'            " AND CONVERT(varchar(8), C.DCTASALDOFECHA, 112) =  ( SELECT MAX(CONVERT(VARCHAR(8), a.dCtaSaldoFecha, 112)) FROM CtaSaldo a " & _
'            " Where a.cCtaContCod = c.cCtaContCod And CONVERT(VARCHAR(8), a.DCTASALDOFECHA, 112) <='" & Format(pdFecha, "YYYYMMdd") & "') " & _
'            " Group by SUBSTRING(C.cCtaContCod, 3,1)"
'    Set oConLocal = New DConecta
'    oConLocal.AbreConexion
'    Set regTemp = oConLocal.CargaRecordSet(lsSql)
'    Do While Not regTemp.EOF
'        If regTemp!nMoneda = "3" Then
'            matFinMes("1", regTemp!Tipo) = matFinMes("1", regTemp!Tipo) + regTemp!nSaldoMN
'        Else
'            matFinMes(regTemp!nMoneda, regTemp!Tipo) = regTemp!nSaldoMN
'        End If
'        regTemp.MoveNext
'    Loop
'    regTemp.Close
'    Set regTemp = Nothing
'    oConLocal.CierraConexion
'
'    lsSql = " select nDesde, nHasta, cDescrip from anxriesgosrango where copecod='770030'"
'    Set oConLocal = New DConecta
'    oConLocal.AbreConexion
'    Set rsRang = oConLocal.CargaRecordSet(lsSql)
'
'    If Not (rsRang.BOF And rsRang.EOF) Then
'        rsRang.MoveLast
'        ReDim lsCodRangINI(rsRang.RecordCount)
'        ReDim lsCodRangFIN(rsRang.RecordCount)
'        ReDim lsDesRang(rsRang.RecordCount)
'
'        lnRangos = rsRang.RecordCount
'        rsRang.MoveFirst
'        i = 0
'        CIIUReg = "("
'        Do While Not rsRang.EOF
'                lsDesRang(i) = rsRang!cDescrip
'                lsCodRangINI(i) = FillNum(Str(rsRang!nDesde), 2, "0")
'                lsCodRangFIN(i) = FillNum(Str(rsRang!nHasta), 2, "0")
'                If lsCodRangINI(i) = lsCodRangFIN(i) Then
'                    CIIUReg = CIIUReg & "'" & lsCodRangINI(i) & "',"
'                Else
'                    For j = lsCodRangINI(i) To lsCodRangFIN(i)
'                        CIIUReg = CIIUReg & "'" & FillNum(Str(j), 2, "0") & "',"
'                    Next
'                End If
'            i = i + 1
'            rsRang.MoveNext
'        Loop
'        CIIUReg = Left(CIIUReg, Len(CIIUReg) - 1) & ")"
'    End If
'
'    For i = 0 To lnRangos - 1
'        If i = -1 Then
'            xlHoja1.Cells(i + 13, 1) = lsDesRang(i)
'            If Trim(lsCodRangINI(i)) = Trim(lsCodRangFIN(i)) Then
'                xlHoja1.Cells(i + 13, 2) = "'" & Trim(lsCodRangINI(i))
'            Else
'                xlHoja1.Cells(i + 13, 2) = "'" & Trim(lsCodRangINI(i)) & " a " & Trim(lsCodRangFIN(i))
'            End If
'            nTFilTemp = i + 13
'        Else
'
'            '=================== SALDO ACTUAL =====================
'
'            If gbBitCentral = True Then
'                lsSql = " SELECT Count(CSC.cCtaCod) as Numero, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'                      " SUM(CASE WHEN CCT.cRefinan = 'N' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN 1 ELSE 0 END) AS NumeroDesembNue, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'N' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'N' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'R' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'R' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefDol " & _
'                      " FROM " & sservidorconsolidada & "CreditoSaldoConsol CSC INNER JOIN " & sservidorconsolidada & " CreditoConsolTotal CCT ON CSC.cCtaCod = CCT.cCtaCod " & _
'                      " LEFT JOIN " & sservidorconsolidada & "FuenteIngresoConsol FI ON CCT.cNumFuente = FI.cNumFuente " & _
'                      " WHERE  ( CSC.nPrdEstado IN(" & cVigente & ") or CSC.nPrdEstado IN (" & gColocEstRecVigJud & ",2205) ) " & _
'                      " AND CSC.nSaldoCap > 0 AND CCT.cTpoCredCod like '[12345]%' " & _
'                      " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(i)) & "' AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(i)) & "' " & _
'                      " AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')"
'            Else
'                lsSql = " SELECT Count(CSC.cCodCta) as Numero, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'                      " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                      " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CCT.nMontoDesemb) END ) AS MontoDesembRefDol " & _
'                      " FROM " & sservidorconsolidada & "CreditoSaldoConsol CSC INNER JOIN DBConsolidada.dbo.CreditoConsolTotal CCT ON CSC.cCodCta = CCT.cCodCta " & _
'                      " LEFT JOIN " & sservidorconsolidada & "FuenteIngresoConsol FI ON CCT.cNumFuente = FI.cNumFuente " & _
'                      " WHERE  ( CSC.cEstado ='F' or (CSC.cEstado = 'V' And CCT.cCondCre = 'J') ) " & _
'                      " AND CSC.nSaldoCap > 0 AND CSC.cCodcta like '__[12]%' " & _
'                      " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(i)) & "' AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(i)) & "' " & _
'                      " AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')"
'            End If
'
'            Set reg9 = oCon.CargaRecordSet(lsSql)
'
'            lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'            lnSaldoSol = IIf(IsNull(reg9!SaldoCapSol), 0, reg9!SaldoCapSol)
'            lnSaldoDol = IIf(IsNull(reg9!SaldoCapDol), 0, reg9!SaldoCapDol * pnTipCambio)
'            lnNumeroDesembNue = IIf(IsNull(reg9!NumeroDesembNue), 0, reg9!NumeroDesembNue)
'            lnDesembNueSol = IIf(IsNull(reg9!MontoDesembNueSol), 0, reg9!MontoDesembNueSol)
'            lnDesembNueDol = IIf(IsNull(reg9!MontoDesembNueDol), 0, reg9!MontoDesembNueDol * pnTipCambio)
'            lnDesembRefSol = IIf(IsNull(reg9!MontoDesembRefSol), 0, reg9!MontoDesembRefSol)
'            lnDesembRefDol = IIf(IsNull(reg9!MontoDesembRefDol), 0, reg9!MontoDesembRefDol * pnTipCambio)
'
'            reg9.Close
'
'            '=================== SALDO MES ANTERIOR =====================
'
'            If gbBitCentral = True Then
'                lsSql = " SELECT Count(CS.cCtaCod) as Numero, " & _
'                      " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'                      " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'                      " FROM " & sservidorconsolidada & "CreditoSaldoConsol CS LEFT JOIN ( SELECT cCtaCod, cTpoCredCod, Max(cNumFuente) cNumFuente FROM " & sservidorconsolidada & "CreditoConsolTotal ct GROUP BY cCtaCod, cTpoCredCod ) c on CS.cCtaCod = C.cCtaCod " & _
'                      " LEFT JOIN  " & sservidorconsolidada & "FuenteIngresoConsol FI  ON C.cNumFuente = FI.cNumFuente" & _
'                      " WHERE CONVERT(VARCHAR(8), CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'                      " AND CS.nPrdEstado IN(" & cVigente & ", " & gColocEstRecVigJud & ",2205) " & _
'                      " AND CS.nSaldoCap > 0 AND C.cTpoCredCod like '[12345]%'   " & _
'                      " AND Substring(CASE WHEN FI.cActEcon IS NULL THEN '5211' ELSE FI.cActEcon END,1,2) >= '" & Trim(lsCodRangINI(i)) & "'" & _
'                      " AND Substring(CASE WHEN FI.cActEcon IS NULL THEN '5211' ELSE FI.cActEcon END,1,2) <= '" & Trim(lsCodRangFIN(i)) & "'"
'            Else
'                lsSql = " SELECT Count(CS.cCodCta) as Numero, " & _
'                      " SUM( CASE WHEN substring(CS.cCodCta,6,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'                      " SUM( CASE WHEN substring(CS.cCodcta,6,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'                      " FROM " & sservidorconsolidada & "CreditoSaldoConsol CS JOIN ( SELECT cCodCta, Max(cNumFuente) cNumFuente FROM " & sservidorconsolidada & "CreditoConsolTotal ct GROUP BY cCodCta ) c on CS.ccodCta = C.cCodCta " & _
'                      " LEFT JOIN " & sservidorconsolidada & "FuenteIngresoConsol FI  ON C.cNumFuente = FI.cNumFuente" & _
'                      " WHERE CONVERT(VARCHAR(8), CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'                      " AND CS.cEstado in('F','V') " & _
'                      " AND CS.nSaldoCap > 0 AND CS.cCodcta like '__[12]%'   " & _
'                      " AND Substring(FI.cActEcon,1,2) >= '" & Trim(lsCodRangINI(i)) & "'" & _
'                      " AND Substring(FI.cActEcon,1,2) <= '" & Trim(lsCodRangFIN(i)) & "'"
'            End If
'            'Revisar este Query
'            Set reg9 = oCon.CargaRecordSet(lsSql)
'            lnSaldoMesAntSol = IIf(IsNull(reg9!SaldoCapAntSol), 0, reg9!SaldoCapAntSol)
'            lnSaldoMesAntDol = IIf(IsNull(reg9!SaldoCapAntDol), 0, reg9!SaldoCapAntDol * lnTipCambMesAnt)
'
'            reg9.Close
'
'            xlHoja1.Cells(i + 13, 1) = lsDesRang(i)
'
'            If Trim(lsCodRangINI(i)) = Trim(lsCodRangFIN(i)) Then
'                xlHoja1.Cells(i + 13, 2) = "'" & Trim(lsCodRangINI(i))
'            Else
'                xlHoja1.Cells(i + 13, 2) = "'" & Trim(lsCodRangINI(i)) & " a " & Trim(lsCodRangFIN(i))
'            End If
'
'            xlHoja1.Cells(i + 13, 3) = lnNroDeudores
'
'            If lsCodRangINI(i) = "01-" And lsCodRangFIN(i) = "02-" Then 'Agricultura
'
'            Else    'Diferente a Agricultura
'                xlHoja1.Cells(i + 13, 4) = lnSaldoSol
'                xlHoja1.Cells(i + 13, 5) = lnSaldoDol
'                xlHoja1.Cells(i + 13, 6) = lnSaldoSol + lnSaldoDol
'
'                If (Mid(lsDesRang(i), 1, 1) <> "-") Then
'                    nTotalTemp(0) = nTotalTemp(0) + lnNroDeudores
'                    nTotalTemp(1) = nTotalTemp(1) + lnSaldoMesAntSol
'                    nTotalTemp(2) = nTotalTemp(2) + lnSaldoMesAntDol
'                    nTotalTemp(3) = nTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'                    nTotalTemp(4) = nTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'                    nTotalTemp(5) = nTotalTemp(5) + lnSaldoSol
'                    nTotalTemp(6) = nTotalTemp(6) + lnSaldoDol
'                    nTotalTemp(7) = nTotalTemp(7) + lnDesembNueSol
'                    nTotalTemp(8) = nTotalTemp(8) + lnDesembNueDol
'                    nTotalTemp(9) = nTotalTemp(9) + lnNumeroDesembNue
'                End If
'
'                If (lsCodRangINI(i) = "50" And lsCodRangFIN(i) = "50") Or (lsCodRangINI(i) = "51" And lsCodRangFIN(i) = "51") Then
'                    nTTotalTemp(0) = nTTotalTemp(0) + lnNroDeudores
'                    nTTotalTemp(1) = nTTotalTemp(1) + lnSaldoMesAntSol
'                    nTTotalTemp(2) = nTTotalTemp(2) + lnSaldoMesAntDol
'                    nTTotalTemp(3) = nTTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'                    nTTotalTemp(4) = nTTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'                    nTTotalTemp(5) = nTTotalTemp(5) + lnSaldoSol
'                    nTTotalTemp(6) = nTTotalTemp(6) + lnSaldoDol
'                    nTTotalTemp(7) = nTTotalTemp(7) + lnDesembNueSol
'                    nTTotalTemp(8) = nTTotalTemp(8) + lnDesembNueDol
'                    nTTotalTemp(9) = nTTotalTemp(9) + lnNumeroDesembNue
'                End If
'            End If
'
'            xlHoja1.Cells(i + 13, 7) = lnNumeroDesembNue
'            xlHoja1.Cells(i + 13, 8) = lnDesembNueSol
'            xlHoja1.Cells(i + 13, 9) = lnDesembNueDol
'
'        End If
'    Next i
'
'   i = i + 1
'
'Dim nOrden As Integer
'
'For lnProduc = 1 To 4
'    '=================== SALDO ACTUAL =====================
'
'        If gbBitCentral = True Then
'            lsSql = " SELECT Count(CC.cCtaCod) as Numero, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'                    " SUM(CASE WHEN (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN 1 ELSE 0 END) AS NumeroDesembNue, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CC.nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CC.nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CC.nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCtaCod,9,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (CC.nMontoDesemb) END ) AS MontoDesembRefDol " & _
'                    " FROM " & sservidorconsolidada & "CreditoConsol CC INNER JOIN " & sservidorconsolidada & "CreditoSaldoConsol CSC ON CC.cCtaCod = CSC.cCtaCod " & _
'                    " INNER JOIN " & sservidorconsolidada & "CreditoConsolTotal CCT ON CSC.cCtaCod = CCT.cCtaCod " & _
'                    " WHERE  ( CSC.nPrdEstado in(" & cVigente & ", " & cPigno & ") or CSC.nPrdEstado IN (" & gColocEstRecVigJud & ",2205) ) " & _
'                    " AND (CSC.nSaldoCap > 0) AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "') "
'        Else
'            lsSql = " SELECT Count(CSC.cCodCta) as Numero, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' THEN CSC.nSaldoCap END ) AS SaldoCapSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' THEN CSC.nSaldoCap END ) AS SaldoCapDol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'N'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembNueDol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodCta,6,1) = '1' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefSol, " & _
'                    " SUM( CASE WHEN substring(CSC.cCodcta,6,1) = '2' AND CCT.cRefinan = 'R'  And (convert(varchar(8), CCT.dFecVig, 112) " & _
'                    " Between '" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "' AND '" & Format(pdFecha, "YYYYmmdd") & "') THEN (nMontoDesemb) END ) AS MontoDesembRefDol " & _
'                    " FROM " & sservidorconsolidada & "CreditoSaldoConsol CSC INNER JOIN " & sservidorconsolidada & "CreditoConsolTotal CCT ON CSC.cCodCta = CCT.cCodCta " & _
'                    " WHERE  ( CSC.cEstado in('F','1','4','6','7') or (CSC.cEstado = 'V' And CCT.cCondCre = 'J') ) " & _
'                    " AND (CSC.nSaldoCap > 0) AND (CONVERT(VARCHAR(8), CSC.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "') "
'        End If
'
'        If gbBitCentral = True Then
'            If lnProduc = 1 Then 'Agricultura
'                lsSql = lsSql & " AND  CCT.cTpoCredCod like '[12345]52' "
'            ElseIf lnProduc = 2 Then  'Corporativo, Grande, Mediana, PEqueña y Micro empresa
'                lsSql = lsSql & " AND  CCT.cTpoCredCod like '[12345]%' "
'            ElseIf lnProduc = 4 Then  ' Consumo  - menos Prendario
'                lsSql = lsSql & " AND  CCT.cTpoCredCod like '7%' --AND CCT.cTpoCredCod <> '755'"
'            ElseIf lnProduc = 3 Then  ' Hipotecarios
'                lsSql = lsSql & " AND  CCT.cTpoCredCod like '8%' "
'            End If
'        Else
'            If lnProduc = 1 Then 'Agricultura
'                lsSql = lsSql & " AND  CSC.cCodcta like '__202%' "
'            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'                lsSql = lsSql & " AND  CSC.cCodcta like '__[12]%' "
'            ElseIf lnProduc = 4 Then  ' Consumo
'                lsSql = lsSql & " AND  CSC.cCodcta like '__3%' "
'            ElseIf lnProduc = 3 Then  ' Hipotecarios
'                lsSql = lsSql & " AND  CSC.cCodcta like '__4%' "
'            End If
'        End If
'
'    Set reg9 = oCon.CargaRecordSet(lsSql)
'
'    'lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'    lnSaldoSol = IIf(IsNull(reg9!SaldoCapSol), 0, reg9!SaldoCapSol)
'    lnSaldoDol = IIf(IsNull(reg9!SaldoCapDol), 0, reg9!SaldoCapDol * pnTipCambio)
'    lnNumeroDesembNue = IIf(IsNull(reg9!NumeroDesembNue), 0, reg9!NumeroDesembNue)
'    lnDesembNueSol = IIf(IsNull(reg9!MontoDesembNueSol), 0, reg9!MontoDesembNueSol)
'    lnDesembNueDol = IIf(IsNull(reg9!MontoDesembNueDol), 0, reg9!MontoDesembNueDol * pnTipCambio)
'    lnDesembRefSol = IIf(IsNull(reg9!MontoDesembRefSol), 0, reg9!MontoDesembRefSol)
'    lnDesembRefDol = IIf(IsNull(reg9!MontoDesembRefDol), 0, reg9!MontoDesembRefDol * pnTipCambio)
'
'    reg9.Close
'    '=================== Deudores por Tipo de Credito =============================
'
'    lsSql = " SELECT Sum(A.Numero) Numero FROM( " & _
'            " SELECT X.Tipo, count(X.cPersCod) Numero from ( " & _
'            " SELECT DISTINCT " & _
'            " Tipo = CASE WHEN C.cTpoCredCod = '755' THEN C.cTpoCredCod ELSE substring(C.cTpoCredCod,1,1) END, P.cPersCod " & _
'            " FROM " & sservidorconsolidada & "CreditoSaldoConsol S " & _
'            " inner join " & sservidorconsolidada & "CreditoConsolTotal C ON S.cCtaCod = C.cCtaCod " & _
'            " inner join " & sservidorconsolidada & "ProductoPersonaConsol P ON P.cCtaCod = C.cCtaCod AND P.nPrdPersRelac = 20 " & _
'            " WHERE  (S.nPrdEstado in (" & cVigente & ", " & cPigno & ") or S.nPrdEstado IN (" & gColocEstRecVigJud & ",2205)) " & _
'            " and (CONVERT(VARCHAR(8), S.dFecha,112)='" & Format(pdFecha, "YYYYmmdd") & "')"
'
'            If lnProduc = 1 Then 'Agricultura
'                lsSql = lsSql & " AND  C.cTpoCredCod like '[12345]52' "
'            ElseIf lnProduc = 2 Then  'Corporativo, Grande, Mediana, PEqueña y Micro empresa
'                lsSql = lsSql & " AND  C.cTpoCredCod like '[12345]%' "
'            ElseIf lnProduc = 4 Then  ' Consumo
'                lsSql = lsSql & " AND  C.cTpoCredCod like '7%'"
'            ElseIf lnProduc = 3 Then  ' Hipotecarios
'                lsSql = lsSql & " AND  C.cTpoCredCod like '8%' "
'            End If
'
'    lsSql = lsSql & " )X GROUP BY X.Tipo) A "
'
'    Set reg9 = oCon.CargaRecordSet(lsSql)
'    lnNroDeudores = IIf(IsNull(reg9!Numero), 0, reg9!Numero)
'    reg9.Close
'    '=================== SALDO MES ANTERIOR =====================
'
'    If gbBitCentral = True Then
'        lsSql = " SELECT Count(c.cCtaCod) as Numero, " & _
'              " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'              " SUM( CASE WHEN substring(CS.cCtaCod,9,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'              " FROM " & sservidorconsolidada & "CreditoSaldoConsol CS LEFT JOIN ( SELECT cCtaCod, Max(cNumFuente) cNumFuente, cTpoCredCod FROM " & sservidorconsolidada & "CreditoConsolTotal ct GROUP BY cCtaCod, cTpoCredCod ) C on CS.cCtaCod = C.cCtaCod " & _
'              " WHERE convert(varchar(8),CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'              " AND CS.nPrdEstado in(" & cVigente & ", " & cPigno & ", " & gColocEstRecVigJud & ",2205) " & _
'              " AND CS.nSaldoCap > 0 "
'
'    Else
'        lsSql = " SELECT Count(c.cCodCta) as Numero, " & _
'              " SUM( CASE WHEN substring(CS.cCodCta,6,1) = '1' THEN CS.nSaldoCap END ) AS SaldoCapAntSol, " & _
'              " SUM( CASE WHEN substring(CS.cCodcta,6,1) = '2' THEN CS.nSaldoCap END ) AS SaldoCapAntDol " & _
'              " FROM " & sservidorconsolidada & "CreditoSaldoConsol CS JOIN ( SELECT cCodCta, Max(cNumFuente) cNumFuente FROM DBConsolidada.dbo.CreditoConsolTotal ct GROUP BY cCodCta ) c on CS.ccodCta = C.cCodCta " & _
'              " WHERE convert(varchar(8),CS.dfecha,112)='" & Format(ldFechaMesAnt, "YYYYMMdd") & "' " & _
'              " AND CS.cEstado in('F','V','1','4','6','7') " & _
'              " AND CS.nSaldoCap > 0 "
'    End If
'
'        If gbBitCentral = True Then
'            If lnProduc = 1 Then 'Agricultura
'                lsSql = lsSql & " AND  C.cTpoCredCod like '[12345]52' "
'            ElseIf lnProduc = 2 Then  'Corporativo, Grande, Mediana, Pequeña y Micro empresa
'                lsSql = lsSql & " AND  C.cTpoCredCod like '[12345]%' "
'            ElseIf lnProduc = 4 Then  ' Consumo
'                lsSql = lsSql & " AND  C.cTpoCredCod like '7%'"
'            ElseIf lnProduc = 3 Then  ' Hipotecarios
'                lsSql = lsSql & " AND  C.cTpoCredCod like '8%' "
'            End If
'
'        Else
'            If lnProduc = 1 Then 'Agricultura
'                lsSql = lsSql & " AND  CS.cCodcta like '__202%' "
'            ElseIf lnProduc = 2 Then  'Comerciales y PYMES
'                lsSql = lsSql & " AND  CS.cCodcta like '__[12]%' "
'            ElseIf lnProduc = 4 Then  ' Consumo
'                lsSql = lsSql & " AND  CS.cCodcta like '__3%' "
'            ElseIf lnProduc = 3 Then  ' Hipotecarios
'                lsSql = lsSql & " AND  CS.cCodcta like '__4%' "
'            End If
'        End If
'
'    Set reg9 = oCon.CargaRecordSet(lsSql)
'
'    lnSaldoMesAntSol = IIf(IsNull(reg9!SaldoCapAntSol), 0, reg9!SaldoCapAntSol)
'    lnSaldoMesAntDol = IIf(IsNull(reg9!SaldoCapAntDol), 0, reg9!SaldoCapAntDol * lnTipCambMesAnt)
'
'    reg9.Close
'
'    '''If bandera = 1 Then
'
'    If lnProduc = 1 Then
'        nTmp = 0
'        xlHoja1.Cells(nTmp + 13, 1) = "A. Agricultura, Ganaderia, Caza y Silvicultura"
'    ElseIf lnProduc = 2 Then
'        nTmp = -1
'        xlHoja1.Cells(nTmp + 13, 1) = "1. CRÉDITOS CORPORATIVOS, A GRANDES, A MEDIANAS, A PEQUEÑAS Y A MICROEMPRESAS"
'        'xlHoja1.Cells(nTmp + 13, 2) = "'01 a 99"
'        nTempoFila(1) = nTmp + 13
'        xlHoja1.Range("A" & nTmp + 13 & ":S" & nTmp + 13).Font.Bold = True
'        xlHoja1.Range("A12").WrapText = True
'        xlHoja1.Range("A12").ColumnWidth = 50
'    Else
'        nTmp = i
'        xlHoja1.Cells(nTmp + 13, 1) = IIf(lnProduc = 4, "3. CREDITOS DE CONSUMO", "2. CREDITOS HIPOTECARIOS PARA VIVIENDA")
'        xlHoja1.Range("A" & nTmp + 12 & ":I" & nTmp + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
'        xlHoja1.Range("A" & nTmp + 13 & ":S" & nTmp + 13).Font.Bold = True
'        If lnProduc = 4 Then
'            nTempoFila(2) = nTmp + 13
'        Else
'            nTempoFila(3) = nTmp + 13
'        End If
'    End If
'
'    xlHoja1.Cells(nTmp + 13, 3) = lnNroDeudores
'    If lnProduc = 2 Then
'        xlHoja1.Range(xlHoja1.Cells(nTmp + 13, 7), xlHoja1.Cells(nTmp + 13, 7)).Formula = "=G43+G42+G41+G40+G39+G36+G35+G34+G33+G29+G28+G27+G16+G15+G14+G13"
'    End If
'    If lnProduc = 1 Then 'Agricultura
'
'        nTotalTemp(0) = nTotalTemp(0) + lnNroDeudores
'        nTotalTemp(1) = nTotalTemp(1) + lnSaldoMesAntSol
'        nTotalTemp(2) = nTotalTemp(2) + lnSaldoMesAntDol
'        nTotalTemp(3) = nTotalTemp(3) + lnDesembNueSol + lnDesembRefSol
'        nTotalTemp(4) = nTotalTemp(4) + lnDesembNueDol + lnDesembRefDol
'        nTotalTemp(5) = nTotalTemp(5) + matFinMes(1, 2)
'        nTotalTemp(6) = nTotalTemp(6) + matFinMes(2, 2)
'        nTotalTemp(7) = nTotalTemp(7) + lnDesembNueSol
'        nTotalTemp(8) = nTotalTemp(8) + lnDesembNueDol
'
'    ElseIf lnProduc = 2 Then ' Corporativos, grandes, medianas, pequeñas y micro empresas
'
'        xlHoja1.Cells(nTmp + 13, 4) = matFinMes(1, 1)
'        xlHoja1.Cells(nTmp + 13, 5) = matFinMes(2, 1)
'        xlHoja1.Cells(nTmp + 13, 6) = matFinMes(1, 1) + matFinMes(2, 1)
'
'    ElseIf lnProduc = 3 Then 'Hipotecarios
'
'        xlHoja1.Cells(nTmp + 13, 4) = matFinMes(1, 4)
'        xlHoja1.Cells(nTmp + 13, 5) = matFinMes(2, 4)
'        xlHoja1.Cells(nTmp + 13, 6) = matFinMes(1, 4) + matFinMes(2, 4)
'
'    ElseIf lnProduc = 4 Then 'De Consumo
'
'        xlHoja1.Cells(nTmp + 13, 4) = matFinMes(1, 3)
'        xlHoja1.Cells(nTmp + 13, 5) = matFinMes(2, 3)
'        xlHoja1.Cells(nTmp + 13, 6) = matFinMes(1, 3) + matFinMes(2, 3)
'    End If
'
'    xlHoja1.Cells(nTmp + 13, 7) = lnNumeroDesembNue
'    xlHoja1.Cells(nTmp + 13, 8) = lnDesembNueSol
'    xlHoja1.Cells(nTmp + 13, 9) = lnDesembNueDol
'
'    If lnProduc <> 1 Then
'        xlHoja1.Range("A" & nTmp + 13 & ":T" & nTmp + 13).Font.Bold = True
'    End If
'
'    If lnProduc = 1 Or lnProduc = 2 Then
'    Else
'        i = i + 2
'    End If
'
' Next
'
'    xlHoja1.Cells(i + 12, 1) = "TOTAL DE CRÉDITOS"
'    xlHoja1.Range("A" & i + 12).HorizontalAlignment = xlCenter
'    ExcelCuadro xlHoja1, 1, i + 12, 9, i + 12
'
'    For nTemp = 99 To 105
'        xlHoja1.Range(UCase(Chr(nTemp)) & i + 12 & ":" & UCase(Chr(nTemp)) & i + 12).Formula = "=+" & UCase(Chr(nTemp)) & nTempoFila(1) & "+" & UCase(Chr(nTemp)) & nTempoFila(2) & "+" & UCase(Chr(nTemp)) & nTempoFila(3)
'    Next
'
'    xlHoja1.Range("A" & i + 12 & ":I" & i + 12).Font.Bold = True
'
'    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(i + 12, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlHoja1.Range("A7:I" & 12 + i).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
'
'    xlHoja1.Cells(i + 16, 1) = "Periodicidad Mensual"
'    xlHoja1.Cells(i + 18, 1) = "(1) Clasificación industrial uniforme de todas las Actividades económicas. Tercera Revisión. Naciones Unidas. Consignar la actividad económica que genera el mayor valor añadido de la entidad deudora"
'    xlHoja1.Cells(i + 19, 1) = "(2) El total de créditos directos debe coincidir con la suma de las cuentas 1401+1403+1404+1405+1406+1407 del Manual de Contabilidad"
'    xlHoja1.Cells(i + 20, 1) = "(3) El total de créditos indirectos debe corresponder a la suma de los saldos de las cuentas 7101+7102+7103+7104+7105 del Manual de Contabilidad"
'
'    xlHoja1.Range("A" & i + 16 & ":A" & i + 16).Font.Bold = True
'    xlHoja1.Range("A" & i + 16 & ":A" & i + 20).Font.Size = 8
'
'    xlHoja1.Range("D" & i + 22 & ":E" & i + 22).MergeCells = True
'    xlHoja1.Range("I" & i + 22 & ":J" & i + 22).MergeCells = True
'    xlHoja1.Range("O" & i + 22 & ":P" & i + 22).MergeCells = True
'
'    xlHoja1.Range("D" & i + 23 & ":E" & i + 23).MergeCells = True
'    xlHoja1.Range("I" & i + 23 & ":J" & i + 23).MergeCells = True
'    xlHoja1.Range("O" & i + 23 & ":P" & i + 23).MergeCells = True
'
'    xlHoja1.Range("I" & i + 24 & ":J" & i + 24).MergeCells = True
'
'    xlHoja1.Cells(i + 23, 4) = "Gerente General"
'    xlHoja1.Cells(i + 23, 9) = "Contador General"
'    xlHoja1.Cells(i + 24, 9) = "Matricula Nro"
'    xlHoja1.Cells(i + 23, 15) = "Hecho Por"
'
'    xlHoja1.Range("A" & i + 23 & ":T" & i + 24).Font.Bold = True
'    xlHoja1.Range("A" & i + 23 & ":T" & i + 24).Font.Size = 8
'
'    xlHoja1.Range("D" & i + 23 & ":P" & i + 24).HorizontalAlignment = xlCenter
'    xlHoja1.Range("D" & i + 23 & ":P" & i + 24).HorizontalAlignment = xlCenter
'    xlHoja1.Range("B13:B" & i + 24).HorizontalAlignment = xlCenter
'
'    xlHoja1.Range("C11:C" & i + 15).NumberFormat = "#,###,##0"
'    xlHoja1.Range("G11:G" & i + 15).NumberFormat = "#,###,##0"
'
'    If gbBitCentral = True Then
'        oConLocal.CierraConexion
'    Else
'        oCon.CierraConexion
'    End If
'   frmMdiMain.staMain.Panels(1).Text = ""
'   RSClose R
'End Sub
'ALPA 20100826 ************************************************************************************
Private Sub GeneraReporteAnexo3(ByVal pdFecha As Date, ByVal pnTipCambio As Double, psMes As String)   ' Flujo Crediticio por Tipo de Credito
Dim I As Integer
Dim nFila As Integer
Dim nIni  As Integer
Dim lNegativo As Boolean
Dim sConec As String
Dim lsSql As String
Dim rsRang As New ADODB.Recordset
Dim lsCodRangINI() As String * 2
Dim lsCodRangFIN() As String * 2
Dim lsCodRango() As String * 2

Dim lsDesRang() As String

Dim nTempoFila(1 To 3) As Integer

Dim lnRangos As Integer
Dim reg9 As New ADODB.Recordset
Dim lnNroDeudores As Long
Dim lnSaldoMesAntSol As Currency, lnSaldoMesAntDol As Currency
Dim lnSaldoSol As Currency, lnSaldoDol As Currency
Dim lnNumeroDesembNue As Long ' BRGO BASILEA II
Dim lnDesembNueSol As Currency, lnDesembNueDol As Currency
Dim lnDesembRefSol As Currency, lnDesembRefDol As Currency
Dim ldFechaMesAnt As Date
Dim CIIUReg As String
Dim lnTipCambMesAnt As Currency
Dim j As Integer
Dim lnProduc As Integer
'Dim nFil As Integer

Dim matFinMes(2, 4) As Currency
Dim regTemp As New ADODB.Recordset
Dim oConLocal As DConecta
Dim nFilTemp As Integer
Dim nTFilTemp As Integer
Dim nTotalTemp(9) As Currency
Dim nTTotalTemp(9) As Currency
Dim nTmp As Integer
Dim nTemp As Integer
    
   ldFechaMesAnt = DateAdd("d", pdFecha, -1 * Day(pdFecha))
   Dim oTC As New nTipoCambio
   lnTipCambMesAnt = oTC.EmiteTipoCambio(ldFechaMesAnt + 1, TCFijoMes)
    
   CabeceraExcelAnexo3 pdFecha, psMes
       
   If Not oCon.AbreConexion Then 'Remota(Right(gsCodAge, 2), True, False, "03")
      Exit Sub
   End If

    lsSql = " select cCodRango, nDesde, nHasta, cDescrip from anxriesgosrango where copecod='770030'"
    Set oConLocal = New DConecta
    oConLocal.AbreConexion
    Set rsRang = oConLocal.CargaRecordSet(lsSql)
      
    If Not (rsRang.BOF And rsRang.EOF) Then
        rsRang.MoveLast
        ReDim lsCodRangINI(rsRang.RecordCount)
        ReDim lsCodRangFIN(rsRang.RecordCount)
        ReDim lsCodRango(rsRang.RecordCount)
        ReDim lsDesRang(rsRang.RecordCount)
        
        lnRangos = rsRang.RecordCount
        rsRang.MoveFirst
        I = 0
        CIIUReg = "("
        Do While Not rsRang.EOF
                lsDesRang(I) = rsRang!cDescrip
                lsCodRangINI(I) = FillNum(Str(rsRang!nDesde), 2, "0")
                lsCodRangFIN(I) = FillNum(Str(rsRang!nHasta), 2, "0")
                lsCodRango(I) = FillNum(rsRang!cCodRango, 2, "0")
                
                If lsCodRangINI(I) = lsCodRangFIN(I) Then
                    CIIUReg = CIIUReg & "'" & lsCodRangINI(I) & "',"
                Else
                    For j = lsCodRangINI(I) To lsCodRangFIN(I)
                        CIIUReg = CIIUReg & "'" & FillNum(Str(j), 2, "0") & "',"
                    Next
                End If
            I = I + 1
            rsRang.MoveNext
        Loop
        CIIUReg = Left(CIIUReg, Len(CIIUReg) - 1) & ")"
    End If
 
    For I = 0 To lnRangos - 1
        If I = -1 Then
            xlHoja1.Cells(I + 14, 1) = lsDesRang(I)
            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I))
            Else
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
            End If
            nTFilTemp = I + 14
        Else
            xlHoja1.Cells(I + 14, 1) = lsDesRang(I)

            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I))
            Else
                If Trim(lsCodRangINI(I)) = "01" Or Trim(lsCodRango(I)) = "05" Or Trim(lsCodRangINI(I)) = "27" Or Trim(lsCodRangINI(I)) = "34" Or Trim(lsCodRangINI(I)) = "40" Or (Trim(lsCodRangINI(I)) = "70" And Trim(lsCodRangFIN(I)) = "71") Or Trim(lsCodRangINI(I)) = "95" Then
                    xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " y " & Trim(lsCodRangFIN(I))
                ElseIf Trim(lsCodRangINI(I)) = "36" Then
                    xlHoja1.Cells(I + 14, 2) = "'23, " & Trim(lsCodRangINI(I)) & " y " & Trim(lsCodRangFIN(I))
                Else
                    xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
                End If
            End If
        End If
    Next I
    lsSql = "exec stp_sel_ObtenerActEconomica "
    oCon.CargaRecordSet (lsSql)
    lsSql = "exec B2_stp_sel_Anexo3_StockFlujoCrediticio '" & Format(pdFecha, "YYYYmmdd") & "','" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "'," & lnTpoCambio
    Set reg9 = oCon.CargaRecordSet(lsSql)
    Do While Not reg9.EOF
        xlHoja1.Cells(reg9!TipoCIIU, 3) = reg9!numero
        xlHoja1.Cells(reg9!TipoCIIU, 4) = reg9!SaldoCapSol
        xlHoja1.Cells(reg9!TipoCIIU, 5) = reg9!SaldoCapDol
        xlHoja1.Cells(reg9!TipoCIIU, 6) = reg9!SaldoCarteTotal
        xlHoja1.Cells(reg9!TipoCIIU, 7) = reg9!NumeroDesembNue
        xlHoja1.Cells(reg9!TipoCIIU, 8) = reg9!MontoDesembNueSol
        xlHoja1.Cells(reg9!TipoCIIU, 9) = reg9!MontoDesembNueDol
    reg9.MoveNext
    Loop
    
    xlHoja1.Cells(14, 1) = "A. Agricultura, Ganaderia, Caza y Silvicultura"
    xlHoja1.Range("A12:I12").Font.Bold = True
    xlHoja1.Cells(12, 1) = "1. CRÉDITOS CORPORATIVOS, A GRANDES, A MEDIANAS, A PEQUEÑAS Y A MICROEMPRESAS"
    xlHoja1.Range("A46:I46").Font.Bold = True
    xlHoja1.Cells(46, 1) = "2. CREDITOS HIPOTECARIOS PARA VIVIENDA"
    xlHoja1.Range("A48:I48").Font.Bold = True
    xlHoja1.Cells(48, 1) = "3. CREDITOS DE CONSUMO"
    xlHoja1.Cells(52, 1) = "TOTAL"
    xlHoja1.Range("A52").HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 1, 52, 9, 52
    
    xlHoja1.Range("C12:C12").Formula = "=+C14+C15+C16+C17+C28+C29+C30+C34+C35+C36+C37+C40+C41+C42+C43+C44"
    xlHoja1.Range("C17:C17").Formula = "=SUM(C18:C27)"
    xlHoja1.Range("C30:C30").Formula = "=SUM(C31:C33)"
    xlHoja1.Range("C37:C37").Formula = "=SUM(C38:C39)"
    xlHoja1.Range("C52:C52").Formula = "=+C12+C48+C46"

    xlHoja1.Range("D12:D12").Formula = "=+D14+D15+D16+D17+D28+D29+D30+D34+D35+D36+D37+D40+D41+D42+D43+D44"
    xlHoja1.Range("D17:D17").Formula = "=SUM(D18:D27)"
    xlHoja1.Range("D30:D30").Formula = "=SUM(D31:D33)"
    xlHoja1.Range("D37:D37").Formula = "=SUM(D38:D39)"
    xlHoja1.Range("D52:D52").Formula = "=+D12+D48+D46"
    
    xlHoja1.Range("E12:E12").Formula = "=+E14+E15+E16+E17+E28+E29+E30+E34+E35+E36+E37+E40+E41+E42+E43+E44"
    xlHoja1.Range("E17:E17").Formula = "=SUM(E18:E27)"
    xlHoja1.Range("E30:E30").Formula = "=SUM(E31:E33)"
    xlHoja1.Range("E37:E37").Formula = "=SUM(E38:E39)"
    xlHoja1.Range("E52:E52").Formula = "=+E12+E48+E46"
    
    xlHoja1.Range("F12:F12").Formula = "=+E12+D12"
    xlHoja1.Range("F17:F17").Formula = "=SUM(F18:F27)"
    xlHoja1.Range("F30:F30").Formula = "=SUM(F31:F33)"
    xlHoja1.Range("F37:F37").Formula = "=SUM(F38:F39)"
    xlHoja1.Range("F52:F52").Formula = "=+F12+F48+F46"
    
    xlHoja1.Range("G12:G12").Formula = "=+G14+G15+G16+G17+G28+G29+G30+G34+G35+G36+G37+G40+G41+G42+G43+G44"
    xlHoja1.Range("G17:G17").Formula = "=SUM(G18:G27)"
    xlHoja1.Range("G30:G30").Formula = "=SUM(G31:G33)"
    xlHoja1.Range("G37:G37").Formula = "=SUM(G38:G39)"
    xlHoja1.Range("G52:G52").Formula = "=+G12+G48+G46"
    
    xlHoja1.Range("H12:H12").Formula = "=+H14+H15+H16+H17+H28+H29+H30+H34+H35+H36+H37+H40+H41+H42+H43+H44"
    xlHoja1.Range("H17:H17").Formula = "=SUM(H18:H27)"
    xlHoja1.Range("H30:H30").Formula = "=SUM(H31:H33)"
    xlHoja1.Range("H37:H37").Formula = "=SUM(H38:H39)"
    xlHoja1.Range("H52:H52").Formula = "=+H12+H48+H46"
    
    xlHoja1.Range("I12:I12").Formula = "=+I14+I15+I16+I17+I28+I29+I30+I34+I35+I36+I37+I40+I41+I42+I43+I44"
    xlHoja1.Range("I17:I17").Formula = "=SUM(I18:I27)"
    xlHoja1.Range("I30:I30").Formula = "=SUM(I31:I33)"
    xlHoja1.Range("I37:I37").Formula = "=SUM(I38:I39)"
    xlHoja1.Range("I52:I52").Formula = "=+I12+I48+I46"
    
    xlHoja1.Range("A52:I52").Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(52, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:I52").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
    xlHoja1.Cells(53, 1) = "Periodicidad Mensual"
    xlHoja1.Cells(55, 1) = "(1) Clasificación industrial uniforme de todas las Actividades económicas. Tercera Revisión. Naciones Unidas. Consignar la actividad económica que genera el mayor valor añadido de la entidad deudora"
    xlHoja1.Cells(56, 1) = "(2) El total de créditos directos debe coincidir con la suma de las cuentas 1401+1403+1404+1405+1406+1407 del Manual de Contabilidad"
    xlHoja1.Cells(57, 1) = "(3) El total de créditos indirectos debe corresponder a la suma de los saldos de las cuentas 7101+7102+7103+7104+7105 del Manual de Contabilidad"
      
    xlHoja1.Range("A55:A57").Font.Bold = True
    xlHoja1.Range("A56:A60").Font.Size = 8
      
    xlHoja1.Range("D59:E59").MergeCells = True
    xlHoja1.Range("I59:J59").MergeCells = True
    xlHoja1.Range("O59:P59").MergeCells = True
     
    xlHoja1.Range("D60:E60").MergeCells = True
    xlHoja1.Range("I60:J60").MergeCells = True
    xlHoja1.Range("O60:P60").MergeCells = True
    
    xlHoja1.Range("I61:J61").MergeCells = True
      
    xlHoja1.Cells(60, 4) = "Gerente General"
    xlHoja1.Cells(60, 9) = "Contador General"
    xlHoja1.Cells(61, 9) = "Matricula Nro"
    xlHoja1.Cells(60, 15) = "Hecho Por"
     
    xlHoja1.Range("A60:T61").Font.Bold = True
    xlHoja1.Range("A60:T61").Font.Size = 8
     
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("B13:B61").HorizontalAlignment = xlCenter
     
    xlHoja1.Range("C11:C52").NumberFormat = "#,###,##0"
    xlHoja1.Range("G11:G52").NumberFormat = "#,###,##0"
   
    If gbBitCentral = True Then
        oConLocal.CierraConexion
    Else
        oCon.CierraConexion
    End If
   frmMdiMain.staMain.Panels(1).Text = ""
   RSClose R
End Sub
'************************************************************************************
Private Sub CabeceraExcelAnexo3(ByVal pdFecha As Date, psMes As String)
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = 46
    xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Cells(4, 1) = "EMPRESA: " & gsNomCmac
    xlHoja1.Cells(5, 9) = "Codigo: " & gsCodCMAC
    xlHoja1.Cells(3, 5) = "STOCK Y FLUJO CREDITICIO POR TIPO DE CREDITO Y SECTOR ECONOMICO"
    xlHoja1.Cells(4, 4) = "Al " & Mid(pdFecha, 1, 2) & " de " & Trim(psMes) & " de " & Year(pdFecha)
    '''xlHoja1.Cells(5, 4) = "( En Nuevos Soles )" 'MARG ERS044-2016
    xlHoja1.Cells(5, 4) = "( En " & StrConv(gcPEN_PLURAL, vbProperCase) & " )" 'MARG ERS044-2016
    xlHoja1.Cells(1, 9) = "ANEXO 3"
      
    xlHoja1.Range("C7:F7").Merge
    xlHoja1.Range("D8:F8").Merge
    xlHoja1.Range("G7:I7").Merge
    xlHoja1.Range("H8:I8").Merge
    
    xlHoja1.Range("A1:S9").HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A1:A50").ColumnWidth = 40
    xlHoja1.Range("B1:B50").ColumnWidth = 10
    xlHoja1.Range("C1:C50").ColumnWidth = 10
    xlHoja1.Range("D1:D50").ColumnWidth = 13
    xlHoja1.Range("E1:E50").ColumnWidth = 13
    xlHoja1.Range("F1:F50").ColumnWidth = 13
    xlHoja1.Range("G1:G50").ColumnWidth = 18
    xlHoja1.Range("H1:H50").ColumnWidth = 18
    xlHoja1.Range("I1:I50").ColumnWidth = 18
   
    xlHoja1.Range("B7:B10").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("A7:I10").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    xlHoja1.Range("A7:B10").Borders(xlInsideVertical).LineStyle = xlContinuous
    'xlHoja1.Range("C7:I10").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("C7:I7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D8:F8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("H8:I8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D7:I10").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:I10").HorizontalAlignment = xlHAlignCenter
    
    'xlHoja1.Cells(8, 1) = "Tipo de Credito"
    xlHoja1.Cells(8, 2) = "Division"
    xlHoja1.Cells(9, 2) = "CIIU (*)"
    xlHoja1.Cells(7, 3) = "STOCK AL CIERRE DEL MES"
    xlHoja1.Cells(8, 3) = "Numero"
    xlHoja1.Cells(9, 3) = "de"
    xlHoja1.Cells(10, 3) = "Deudores"
    xlHoja1.Cells(8, 4) = "Saldo"
    xlHoja1.Cells(9, 4) = "M.N."
    '''xlHoja1.Cells(10, 4) = "(Miles de N.S.)" 'MARG ERS044-2016
    xlHoja1.Cells(10, 4) = "(Miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(9, 5) = "M.E."
    '''xlHoja1.Cells(10, 5) = "(Miles de N.S.)" 'MARG ERS044-2016
    xlHoja1.Cells(10, 5) = "(Miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(9, 6) = "Total"
    '''xlHoja1.Cells(10, 6) = "(Miles de N.S.)" 'MARG ERS044-2016
    xlHoja1.Cells(10, 6) = "(Miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    
    xlHoja1.Cells(7, 7) = "FLUJO DESEMBOLSADO EN EL MES"
    xlHoja1.Cells(8, 7) = "Número de nuevos"
    xlHoja1.Cells(9, 7) = "créditos"
    xlHoja1.Cells(10, 7) = "desembolsados"
    xlHoja1.Cells(8, 8) = "Monto de nuevos créditos desemb."
    xlHoja1.Cells(9, 8) = "M.N."
    '''xlHoja1.Cells(10, 8) = "(Miles de N.S.)" 'MARG ERS044-2016
    xlHoja1.Cells(10, 8) = "(Miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(9, 9) = "M.E."
    xlHoja1.Cells(10, 9) = "(Miles de US$)"
    
    xlHoja1.Range("D12:F50").NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range("H12:I50").NumberFormat = "#,##0.00;-#,##0.00"
End Sub


Public Sub GeneraAnx05D(pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String)
                
Dim sql As String
Dim rsP As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim lsHoja As String
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim FechaFinMes As Date
Dim n As Integer
Dim Conecta As DConecta

Dim lsFormat As String
Dim lbLibroOpen As Boolean

Set Conecta = New DConecta
lsFormat = CStr(pnAnio) + CStr(pnMes)

'RaiseEvent ShowProgress
lsArchivo = App.path & "\Spooler\Anx05D_" & lsFormat & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
If lbLibroOpen Then
   For n = 1 To 2
        If n = 1 Then
            lsHoja = "En Miles"
        Else
            lsHoja = "En Nuevos Soles"
        End If
       
       ExcelAddHoja lsHoja, xlLibro, xlHoja1
    '          oBarra.Progress 1, lsTitulo, "Hoja Excel generada...", , vbBlue
    '   Call GeneraReporteAnexo5(pnAnio, pnMes, pnTpoCambio, psMes)
        Conecta.AbreConexion
        sql = "Select * From DbcMactConsolidada.dbo.VarConsolida"
    
        Set rs = Conecta.CargaRecordSet(sql)
        dFecha = rs!dfecConsol
        Set rs = Nothing
    
        xlHoja1.Cells(2, 1) = "ANEXO N°5-D"
        xlHoja1.Range("A2:G2").Font.Bold = True
        xlHoja1.Range("A2:G2").Font.Size = 12
        xlHoja1.Range("A2:G2").HorizontalAlignment = xlCenter
        xlHoja1.Range("A2:G2").MergeCells = True
        xlHoja1.Cells(3, 1) = "EMPRESA : Caja Metropolitana de Lima"
        xlHoja1.Cells(3, 6) = "CODIGO : 102"
        xlHoja1.Range("A3:F3").Font.Bold = True
        xlHoja1.Cells(5, 1) = "INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS QUE RESPALDAN FINANCIAMIENTOS O LÍNEAS DE CRÉDITO"
        xlHoja1.Range("A5:G5").Font.Bold = True
        xlHoja1.Range("A5:G5").Font.Size = 9
        xlHoja1.Range("A5:G5").HorizontalAlignment = xlCenter
        xlHoja1.Range("A5:G5").MergeCells = True
        xlHoja1.Cells(7, 1) = "Al " + Mid(CStr(dFecha), 1, 2) + " de " + Trim(psMes) + " del " + Mid(CStr(dFecha), 7, 4)
        xlHoja1.Range("A7:G7").Font.Bold = True
        xlHoja1.Range("A7:G7").Font.Size = 9
        xlHoja1.Range("A7:G7").HorizontalAlignment = xlCenter
        xlHoja1.Range("A7:G7").MergeCells = True
        If n = 1 Then
            xlHoja1.Cells(8, 1) = "(En Miles de Nuevo Soles)"
        Else
            xlHoja1.Cells(8, 1) = "(En Nuevo Soles)"
        End If
        xlHoja1.Range("A8:G8").Font.Bold = True
        xlHoja1.Range("A8:G8").Font.Size = 9
        xlHoja1.Range("A8:G8").HorizontalAlignment = xlCenter
        xlHoja1.Range("A8:G8").MergeCells = True
        
        xlHoja1.Range("B10:G46").HorizontalAlignment = xlCenter
        xlHoja1.Cells(10, 1) = "A.- MONTO DE LOS CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS"
        xlHoja1.Range("A10:A12").MergeCells = True
        xlHoja1.Range("A10:A12").WrapText = True
        xlHoja1.Range("A10:A12").VerticalAlignment = xlCenter
        xlHoja1.Range("A10:A12").HorizontalAlignment = xlCenter
        xlHoja1.Cells(11, 2) = "Normal"
        xlHoja1.Range("B10:B12").MergeCells = True
        xlHoja1.Range("B10:B12").VerticalAlignment = xlCenter
        xlHoja1.Cells(11, 3) = "CPP"
        xlHoja1.Range("C10:C12").MergeCells = True
        xlHoja1.Range("C10:C12").VerticalAlignment = xlCenter
        xlHoja1.Cells(11, 4) = "Deficiente"
        xlHoja1.Range("D10:D12").MergeCells = True
        xlHoja1.Range("D10:D12").VerticalAlignment = xlCenter
        xlHoja1.Cells(11, 5) = "Dudoso"
        xlHoja1.Range("E10:E12").MergeCells = True
        xlHoja1.Range("E10:E12").VerticalAlignment = xlCenter
        xlHoja1.Cells(11, 6) = "Perdida"
        xlHoja1.Range("F10:F12").MergeCells = True
        xlHoja1.Range("F10:F12").VerticalAlignment = xlCenter
        xlHoja1.Cells(11, 7) = "Total"
        xlHoja1.Range("G10:G12").MergeCells = True
        xlHoja1.Range("G10:G12").VerticalAlignment = xlCenter
        xlHoja1.Cells(13, 1) = "Comerciales"
        xlHoja1.Cells(14, 1) = "MES"
        xlHoja1.Cells(15, 1) = "Hipotecario para Vivienda"
        xlHoja1.Cells(16, 1) = "Consumo"
        xlHoja1.Range("A10:G12").Cells.Borders.LineStyle = xlContinuous
        
        xlHoja1.Cells(18, 1) = "B.- PROVISIONES CONSTITUIDAS"
        xlHoja1.Range("A17:A18").MergeCells = True
        xlHoja1.Range("A17:A18").VerticalAlignment = xlCenter
        xlHoja1.Range("A17:A18").HorizontalAlignment = xlCenter
        xlHoja1.Cells(18, 2) = "Normal"
        xlHoja1.Range("B17:B18").MergeCells = True
        xlHoja1.Range("B17:B18").VerticalAlignment = xlCenter
        xlHoja1.Cells(18, 3) = "CPP"
        xlHoja1.Range("C17:C18").MergeCells = True
        xlHoja1.Range("C17:C18").VerticalAlignment = xlCenter
        xlHoja1.Cells(18, 4) = "Deficiente"
        xlHoja1.Range("D17:D18").MergeCells = True
        xlHoja1.Range("D17:D18").VerticalAlignment = xlCenter
        xlHoja1.Cells(18, 5) = "Dudoso"
        xlHoja1.Range("E17:E18").MergeCells = True
        xlHoja1.Range("E17:E18").VerticalAlignment = xlCenter
        xlHoja1.Cells(18, 6) = "Perdida"
        xlHoja1.Range("F17:F18").MergeCells = True
        xlHoja1.Range("F17:F18").VerticalAlignment = xlCenter
        xlHoja1.Cells(18, 7) = "Total"
        xlHoja1.Range("G17:G18").MergeCells = True
        xlHoja1.Range("G17:G18").VerticalAlignment = xlCenter
        xlHoja1.Cells(19, 1) = "Comerciales"
        xlHoja1.Cells(20, 1) = "MES"
        xlHoja1.Cells(21, 1) = "Hipotecario para Vivienda"
        xlHoja1.Cells(22, 1) = "Consumo"
        xlHoja1.Range("A17:G18").Cells.Borders.LineStyle = xlContinuous
        
        xlHoja1.Cells(24, 1) = "C.- PROVISIONES REQUERIDAS"
        xlHoja1.Range("A23:A24").MergeCells = True
        xlHoja1.Range("A23:A24").VerticalAlignment = xlCenter
        xlHoja1.Range("A23:A24").HorizontalAlignment = xlCenter
        xlHoja1.Cells(24, 2) = "Normal"
        xlHoja1.Range("B23:B24").MergeCells = True
        xlHoja1.Range("B23:B24").VerticalAlignment = xlCenter
        xlHoja1.Cells(24, 3) = "CPP"
        xlHoja1.Range("C23:C24").MergeCells = True
        xlHoja1.Range("C23:C24").VerticalAlignment = xlCenter
        xlHoja1.Cells(24, 4) = "Deficiente"
        xlHoja1.Range("D23:D24").MergeCells = True
        xlHoja1.Range("D23:D24").VerticalAlignment = xlCenter
        xlHoja1.Cells(24, 5) = "Dudoso"
        xlHoja1.Range("E23:E24").MergeCells = True
        xlHoja1.Range("E23:E24").VerticalAlignment = xlCenter
        xlHoja1.Cells(24, 6) = "Perdida"
        xlHoja1.Range("F23:F24").MergeCells = True
        xlHoja1.Range("F23:F24").VerticalAlignment = xlCenter
        xlHoja1.Cells(24, 7) = "Total"
        xlHoja1.Range("G23:G24").MergeCells = True
        xlHoja1.Range("G23:G24").VerticalAlignment = xlCenter
        xlHoja1.Cells(25, 1) = "Comerciales"
        xlHoja1.Cells(26, 1) = "MES"
        xlHoja1.Cells(27, 1) = "Hipotecario para Vivienda"
        xlHoja1.Cells(28, 1) = "Consumo"
        xlHoja1.Range("A23:G24").Cells.Borders.LineStyle = xlContinuous
        
        xlHoja1.Cells(29, 1) = "D.- SUPERÁVIT (DÉFICIT) DE PROVISIONES"
        xlHoja1.Range("A29:A30").MergeCells = True
        xlHoja1.Range("A29:A30").VerticalAlignment = xlCenter
        xlHoja1.Range("A29:A30").HorizontalAlignment = xlCenter
        xlHoja1.Range("A29:A30").WrapText = True
        xlHoja1.Cells(30, 2) = "Normal"
        xlHoja1.Range("B29:B30").MergeCells = True
        xlHoja1.Range("B29:B30").VerticalAlignment = xlCenter
        xlHoja1.Cells(30, 3) = "CPP"
        xlHoja1.Range("C29:C30").MergeCells = True
        xlHoja1.Range("C29:C30").VerticalAlignment = xlCenter
        xlHoja1.Cells(30, 4) = "Deficiente"
        xlHoja1.Range("D29:D30").MergeCells = True
        xlHoja1.Range("D29:D30").VerticalAlignment = xlCenter
        xlHoja1.Cells(30, 5) = "Dudoso"
        xlHoja1.Range("E29:E30").MergeCells = True
        xlHoja1.Range("E29:E30").VerticalAlignment = xlCenter
        xlHoja1.Cells(30, 6) = "Perdida"
        xlHoja1.Range("F29:F30").MergeCells = True
        xlHoja1.Range("F29:F30").VerticalAlignment = xlCenter
        xlHoja1.Cells(30, 7) = "Total"
        xlHoja1.Range("G29:G30").MergeCells = True
        xlHoja1.Range("G29:G30").VerticalAlignment = xlCenter
        xlHoja1.Cells(31, 1) = "Comerciales"
        xlHoja1.Cells(32, 1) = "MES"
        xlHoja1.Cells(33, 1) = "Hipotecario para Vivienda"
        xlHoja1.Cells(34, 1) = "Consumo"
        xlHoja1.Cells(35, 1) = "TOTAL"
        xlHoja1.Range("A29:G30").Cells.Borders.LineStyle = xlContinuous
        
        xlHoja1.Cells(36, 1) = "E.- MONTO DE LOS CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS CEDIDOS EN EJECUCION"
        xlHoja1.Range("A36:A40").MergeCells = True
        xlHoja1.Range("A36:A40").VerticalAlignment = xlCenter
        xlHoja1.Range("A36:A40").HorizontalAlignment = xlCenter
        xlHoja1.Range("A36:A40").WrapText = True
        xlHoja1.Cells(36, 2) = "Normal"
        xlHoja1.Range("B36:B40").MergeCells = True
        xlHoja1.Range("B36:B40").VerticalAlignment = xlCenter
        xlHoja1.Cells(36, 3) = "CPP"
        xlHoja1.Range("C36:C40").MergeCells = True
        xlHoja1.Range("C36:C40").VerticalAlignment = xlCenter
        xlHoja1.Cells(36, 4) = "Deficiente"
        xlHoja1.Range("D36:D40").MergeCells = True
        xlHoja1.Range("D36:D40").VerticalAlignment = xlCenter
        xlHoja1.Cells(36, 5) = "Dudoso"
        xlHoja1.Range("E36:E40").MergeCells = True
        xlHoja1.Range("E36:E40").VerticalAlignment = xlCenter
        xlHoja1.Cells(36, 6) = "Perdida"
        xlHoja1.Range("F36:F40").MergeCells = True
        xlHoja1.Range("F36:F40").VerticalAlignment = xlCenter
        xlHoja1.Cells(36, 7) = "Total"
        xlHoja1.Range("G36:G40").MergeCells = True
        xlHoja1.Range("G36:G40").VerticalAlignment = xlCenter
        xlHoja1.Cells(41, 1) = "Comerciales"
        xlHoja1.Cells(42, 1) = "MES"
        xlHoja1.Cells(43, 1) = "Hipotecario para Vivienda"
        xlHoja1.Cells(44, 1) = "Consumo"
        xlHoja1.Range("A36:G40").Cells.Borders.LineStyle = xlContinuous
        
        xlHoja1.Cells(48, 2).Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Cells(48, 2) = "GERENTE"
        xlHoja1.Cells(48, 2).HorizontalAlignment = xlCenter
        xlHoja1.Cells(49, 2) = "GENERAL"
        xlHoja1.Cells(49, 2).HorizontalAlignment = xlCenter
        xlHoja1.Cells(48, 4).Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Cells(48, 4) = "CONTADOR"
        xlHoja1.Cells(48, 4).HorizontalAlignment = xlCenter
        xlHoja1.Cells(49, 4) = "GENERAL"
        xlHoja1.Cells(49, 4).HorizontalAlignment = xlCenter
        xlHoja1.Cells(48, 6).Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Cells(48, 6) = "FUNCIONARIO"
        xlHoja1.Cells(48, 6).HorizontalAlignment = xlCenter
        xlHoja1.Cells(49, 6) = "RESPONSABLE"
        xlHoja1.Cells(49, 6).HorizontalAlignment = xlCenter
        
        xlHoja1.Range("A13:A44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("B13:B44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("C13:C44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("D13:D44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("E13:E44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("F13:F44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("G13:G44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("H13:H44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("A35:G35").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("A45:G45").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
        
        xlHoja1.Range("B13:G16") = Format(0, "##,###,###,##0")
        xlHoja1.Range("B19:G22") = Format(0, "##,###,###,##0")
        xlHoja1.Range("B25:G28") = Format(0, "##,###,###,##0")
        xlHoja1.Range("B31:G35") = Format(0, "##,###,###,##0")
        xlHoja1.Range("B41:G44") = Format(0, "##,###,###,##0")
        
        sql = "Select nNormal nSaldo01, nCpp nSaldo02, nDeficiente nSaldo03, nDudoso nSaldo04, nPerdida nSaldo05 "
        sql = sql + " From Anexo5Resumen "
        sql = sql + " Where nAño = " & pnAnio & " And nMes = " & pnMes
        sql = sql + " And cSubGrupo='A' And nSecuencia = 13"
        Set rsP = Conecta.CargaRecordSet(sql)
        While Not rsP.EOF
            If n = 1 Then
                xlHoja1.Cells(15, 2) = Format(rsP!nSaldo01 / 1000, "##,###,###,##0")
                xlHoja1.Cells(15, 3) = Format(rsP!nSaldo02 / 1000, "##,###,###,##0")
                xlHoja1.Cells(15, 4) = Format(rsP!nSaldo03 / 1000, "##,###,###,##0")
                xlHoja1.Cells(15, 5) = Format(rsP!nSaldo04 / 1000, "##,###,###,##0")
                xlHoja1.Cells(15, 6) = Format(rsP!nSaldo05 / 1000, "##,###,###,##0")
            Else
                xlHoja1.Cells(15, 2) = Format(rsP!nSaldo01, IIf(rsP!nSaldo01 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(15, 3) = Format(rsP!nSaldo02, IIf(rsP!nSaldo02 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(15, 4) = Format(rsP!nSaldo03, IIf(rsP!nSaldo03 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(15, 5) = Format(rsP!nSaldo04, IIf(rsP!nSaldo04 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(15, 6) = Format(rsP!nSaldo05, IIf(rsP!nSaldo05 = 0, "##,###,###,##0", "##,###,###,##0.00"))
            End If
            rsP.MoveNext
        Wend
        Set rsP = Nothing
        
        sql = "Select nNormal nProvision01, nCpp nProvision02, nDeficiente nProvision03, nDudoso nProvision04, nPerdida nProvision05"
        sql = sql + " From Anexo5Resumen"
        sql = sql + " Where nAño = " & pnAnio & " And nMES = " & pnMes
        sql = sql + " And cSubGrupo = 'K' And nSecuencia = 63"
        Set rsP = Conecta.CargaRecordSet(sql)
        While Not rsP.EOF
            If n = 1 Then
                xlHoja1.Cells(21, 2) = Format(rsP!nProvision01 / 1000, "##,###,###,##0")
                xlHoja1.Cells(21, 3) = Format(rsP!nProvision02 / 1000, "##,###,###,##0")
                xlHoja1.Cells(21, 4) = Format(rsP!nProvision03 / 1000, "##,###,###,##0")
                xlHoja1.Cells(21, 5) = Format(rsP!nProvision04 / 1000, "##,###,###,##0")
                xlHoja1.Cells(21, 6) = Format(rsP!nProvision05 / 1000, "##,###,###,##0")
            Else
                xlHoja1.Cells(21, 2) = Format(rsP!nProvision01, IIf(rsP!nProvision01 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(21, 3) = Format(rsP!nProvision02, IIf(rsP!nProvision02 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(21, 4) = Format(rsP!nProvision03, IIf(rsP!nProvision03 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(21, 5) = Format(rsP!nProvision04, IIf(rsP!nProvision04 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(21, 6) = Format(rsP!nProvision05, IIf(rsP!nProvision05 = 0, "##,###,###,##0", "##,###,###,##0.00"))
            End If
            rsP.MoveNext
        Wend
        Set rsP = Nothing
        
        sql = "Select nNormal nProvision01, nCpp nProvision02, nDeficiente nProvision03, nDudoso nProvision04, nPerdida nProvision05"
        sql = sql + " From Anexo5Resumen"
        sql = sql + " Where nAño = " & pnAnio & " And nMES = " & pnMes
        sql = sql + " And cSubGrupo = 'L' And nSecuencia = 68"
        Set rsP = Conecta.CargaRecordSet(sql)
        While Not rsP.EOF
            If n = 1 Then
                xlHoja1.Cells(27, 2) = Format(rsP!nProvision01 / 1000, "##,###,###,##0")
                xlHoja1.Cells(27, 3) = Format(rsP!nProvision02 / 1000, "##,###,###,##0")
                xlHoja1.Cells(27, 4) = Format(rsP!nProvision03 / 1000, "##,###,###,##0")
                xlHoja1.Cells(27, 5) = Format(rsP!nProvision04 / 1000, "##,###,###,##0")
                xlHoja1.Cells(27, 6) = Format(rsP!nProvision05 / 1000, "##,###,###,##0")
            Else
                xlHoja1.Cells(27, 2) = Format(rsP!nProvision01, IIf(rsP!nProvision01 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(27, 3) = Format(rsP!nProvision02, IIf(rsP!nProvision02 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(27, 4) = Format(rsP!nProvision03, IIf(rsP!nProvision03 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(27, 5) = Format(rsP!nProvision04, IIf(rsP!nProvision04 = 0, "##,###,###,##0", "##,###,###,##0.00"))
                xlHoja1.Cells(27, 6) = Format(rsP!nProvision05, IIf(rsP!nProvision05 = 0, "##,###,###,##0", "##,###,###,##0.00"))
            End If
            rsP.MoveNext
        Wend
        Conecta.CierraConexion
        rsP.Close
        Set rsP = Nothing
        xlHoja1.Cells(15, 7).Formula = "=SUM(B15:F15)"
        xlHoja1.Cells(21, 7).Formula = "=SUM(B21:F21)"
        xlHoja1.Cells(27, 7).Formula = "=SUM(B27:F27)"
        xlHoja1.Range("A1:A50").ColumnWidth = 33
        xlHoja1.Range("B1:B50").ColumnWidth = 18
        xlHoja1.Range("C1:C50").ColumnWidth = 18
        xlHoja1.Range("D1:D50").ColumnWidth = 18
        xlHoja1.Range("E1:E50").ColumnWidth = 18
        xlHoja1.Range("F1:F50").ColumnWidth = 18
        xlHoja1.Range("G1:G50").ColumnWidth = 18
        xlHoja1.PageSetup.CenterHorizontally = True
        xlHoja1.PageSetup.Zoom = 85
        xlHoja1.PageSetup.Orientation = xlLandscape
        xlHoja1.Range("A1:G50").Font.Size = 9
        xlHoja1.Range("A2:G2").Font.Size = 12
        xlHoja1.Range("A5:G5").Font.Size = 8
   Next
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        MsgBox "Archivo del Anexo5D generado satisfactoriamente", vbInformation, "Aviso!!!"
   CargaArchivo lsArchivo, App.path & "\Spooler"
End If

End Sub

'ALPA 20120118*************************************************************
Public Sub GeneraAnx03FlujoCrediticioPorTipoCredRiesgos(ByVal pnBitCentral As Boolean, pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim n           As Integer
Dim ldFechaRep As Date
 
On Error GoTo ErrImprimeRiesgos
 
MousePointer = 11
lnTpoCambio = pnTpoCambio
 
If pnBandera = 1 Then
    lsArchivo = App.path & "\Spooler\Anx03_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(Str(pnMes)) & gsCodUser, Trim(Str(pnMes))) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        ExcelAddHoja psMes, xlLibro, xlHoja1
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        Call GeneraReporteAnexo3(ldFechaRep, pnTpoCambio, psMes)
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        CargaArchivo lsArchivo, App.path & "\Spooler"
    End If
    MousePointer = 0
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
ElseIf pnBandera = 2 Then
    MousePointer = 11
    
    If pnBitCentral = True Then
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        GeneraSUCAVEAnx03 pnBitCentral, ldFechaRep, psMes
    Else
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        GeneraSUCAVEAnx03 pnBitCentral, ldFechaRep, psMes
    End If

End If

Exit Sub
ErrImprimeRiesgos:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   MousePointer = 0
End Sub

Private Sub GeneraReporteAnexo3Riesgos(ByVal pdFecha As Date, ByVal pnTipCambio As Double, psMes As String)   ' Flujo Crediticio por Tipo de Credito
Dim I As Integer
Dim nFila As Integer
Dim nIni  As Integer
Dim lNegativo As Boolean
Dim sConec As String
Dim lsSql As String
Dim rsRang As New ADODB.Recordset
Dim lsCodRangINI() As String * 2
Dim lsCodRangFIN() As String * 2
Dim lsCodRango() As String * 2

Dim lsDesRang() As String

Dim nTempoFila(1 To 3) As Integer

Dim lnRangos As Integer
Dim reg9 As New ADODB.Recordset
Dim lnNroDeudores As Long
Dim lnSaldoMesAntSol As Currency, lnSaldoMesAntDol As Currency
Dim lnSaldoSol As Currency, lnSaldoDol As Currency
Dim lnNumeroDesembNue As Long ' BRGO BASILEA II
Dim lnDesembNueSol As Currency, lnDesembNueDol As Currency
Dim lnDesembRefSol As Currency, lnDesembRefDol As Currency
Dim ldFechaMesAnt As Date
Dim CIIUReg As String
Dim lnTipCambMesAnt As Currency
Dim j As Integer
Dim lnProduc As Integer
'Dim nFil As Integer

Dim matFinMes(2, 4) As Currency
Dim regTemp As New ADODB.Recordset
Dim oConLocal As DConecta
Dim nFilTemp As Integer
Dim nTFilTemp As Integer
Dim nTotalTemp(9) As Currency
Dim nTTotalTemp(9) As Currency
Dim nTmp As Integer
Dim nTemp As Integer
    
   ldFechaMesAnt = DateAdd("d", pdFecha, -1 * Day(pdFecha))
   Dim oTC As New nTipoCambio
   lnTipCambMesAnt = oTC.EmiteTipoCambio(ldFechaMesAnt + 1, TCFijoMes)
    
   CabeceraExcelAnexo3Riesgos pdFecha, psMes
       
   If Not oCon.AbreConexion Then 'Remota(Right(gsCodAge, 2), True, False, "03")
      Exit Sub
   End If

    lsSql = " select cCodRango, nDesde, nHasta, cDescrip from anxriesgosrango where copecod='770030'"
    Set oConLocal = New DConecta
    oConLocal.AbreConexion
    Set rsRang = oConLocal.CargaRecordSet(lsSql)
      
    If Not (rsRang.BOF And rsRang.EOF) Then
        rsRang.MoveLast
        ReDim lsCodRangINI(rsRang.RecordCount)
        ReDim lsCodRangFIN(rsRang.RecordCount)
        ReDim lsCodRango(rsRang.RecordCount)
        ReDim lsDesRang(rsRang.RecordCount)
        
        lnRangos = rsRang.RecordCount
        rsRang.MoveFirst
        I = 0
        CIIUReg = "("
        Do While Not rsRang.EOF
                lsDesRang(I) = rsRang!cDescrip
                lsCodRangINI(I) = FillNum(Str(rsRang!nDesde), 2, "0")
                lsCodRangFIN(I) = FillNum(Str(rsRang!nHasta), 2, "0")
                lsCodRango(I) = FillNum(rsRang!cCodRango, 2, "0")
                
                If lsCodRangINI(I) = lsCodRangFIN(I) Then
                    CIIUReg = CIIUReg & "'" & lsCodRangINI(I) & "',"
                Else
                    For j = lsCodRangINI(I) To lsCodRangFIN(I)
                        CIIUReg = CIIUReg & "'" & FillNum(Str(j), 2, "0") & "',"
                    Next
                End If
            I = I + 1
            rsRang.MoveNext
        Loop
        CIIUReg = Left(CIIUReg, Len(CIIUReg) - 1) & ")"
    End If
 
    For I = 0 To lnRangos - 1
        If I = -1 Then
            xlHoja1.Cells(I + 14, 1) = lsDesRang(I)
            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I))
            Else
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
            End If
            nTFilTemp = I + 14
        Else
            xlHoja1.Cells(I + 14, 1) = lsDesRang(I)

            If Trim(lsCodRangINI(I)) = Trim(lsCodRangFIN(I)) Then
                xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I))
            Else
                If Trim(lsCodRangINI(I)) = "01" Or Trim(lsCodRango(I)) = "05" Or Trim(lsCodRangINI(I)) = "27" Or Trim(lsCodRangINI(I)) = "34" Or Trim(lsCodRangINI(I)) = "40" Or (Trim(lsCodRangINI(I)) = "70" And Trim(lsCodRangFIN(I)) = "71") Or Trim(lsCodRangINI(I)) = "95" Then
                    xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " y " & Trim(lsCodRangFIN(I))
                ElseIf Trim(lsCodRangINI(I)) = "36" Then
                    xlHoja1.Cells(I + 14, 2) = "'23, " & Trim(lsCodRangINI(I)) & " y " & Trim(lsCodRangFIN(I))
                Else
                    xlHoja1.Cells(I + 14, 2) = "'" & Trim(lsCodRangINI(I)) & " a " & Trim(lsCodRangFIN(I))
                End If
            End If
        End If
    Next I
    lsSql = "exec stp_sel_ObtenerActEconomica "
    oCon.CargaRecordSet (lsSql)
    lsSql = "exec B2_stp_sel_Anexo3_StockFlujoCrediticioRiesgos '" & Format(pdFecha, "YYYYmmdd") & "','" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "'," & lnTpoCambio
    Set reg9 = oCon.CargaRecordSet(lsSql)
    Do While Not reg9.EOF
        xlHoja1.Cells(reg9!TipoCIIU, 3) = reg9!numero
        xlHoja1.Cells(reg9!TipoCIIU, 4) = reg9!SaldoCapSol
        xlHoja1.Cells(reg9!TipoCIIU, 5) = reg9!SaldoCapDol
        xlHoja1.Cells(reg9!TipoCIIU, 6) = reg9!SaldoCarteTotal
        xlHoja1.Cells(reg9!TipoCIIU, 7) = reg9!NumeroDesembNue
        xlHoja1.Cells(reg9!TipoCIIU, 8) = reg9!MontoDesembNueSol
        xlHoja1.Cells(reg9!TipoCIIU, 9) = reg9!MontoDesembNueDol
    reg9.MoveNext
    Loop
    
    xlHoja1.Cells(14, 1) = "A. Agricultura, Ganaderia, Caza y Silvicultura"
    xlHoja1.Range("A12:I12").Font.Bold = True
    xlHoja1.Cells(12, 1) = "1. CRÉDITOS CORPORATIVOS, A GRANDES, A MEDIANAS, A PEQUEÑAS Y A MICROEMPRESAS"
    xlHoja1.Range("A46:I46").Font.Bold = True
    xlHoja1.Cells(46, 1) = "2. CREDITOS HIPOTECARIOS PARA VIVIENDA"
    xlHoja1.Range("A48:I48").Font.Bold = True
    xlHoja1.Cells(48, 1) = "3. CREDITOS DE CONSUMO"
    xlHoja1.Cells(52, 1) = "TOTAL"
    xlHoja1.Range("A52").HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 1, 52, 9, 52
    
    xlHoja1.Range("C12:C12").Formula = "=+C14+C15+C16+C17+C28+C29+C30+C34+C35+C36+C37+C40+C41+C42+C43+C44"
    xlHoja1.Range("C17:C17").Formula = "=SUM(C18:C27)"
    xlHoja1.Range("C30:C30").Formula = "=SUM(C31:C33)"
    xlHoja1.Range("C37:C37").Formula = "=SUM(C38:C39)"
    xlHoja1.Range("C52:C52").Formula = "=+C12+C48+C46"

    xlHoja1.Range("D12:D12").Formula = "=+D14+D15+D16+D17+D28+D29+D30+D34+D35+D36+D37+D40+D41+D42+D43+D44"
    xlHoja1.Range("D17:D17").Formula = "=SUM(D18:D27)"
    xlHoja1.Range("D30:D30").Formula = "=SUM(D31:D33)"
    xlHoja1.Range("D37:D37").Formula = "=SUM(D38:D39)"
    xlHoja1.Range("D52:D52").Formula = "=+D12+D48+D46"
    
    xlHoja1.Range("E12:E12").Formula = "=+E14+E15+E16+E17+E28+E29+E30+E34+E35+E36+E37+E40+E41+E42+E43+E44"
    xlHoja1.Range("E17:E17").Formula = "=SUM(E18:E27)"
    xlHoja1.Range("E30:E30").Formula = "=SUM(E31:E33)"
    xlHoja1.Range("E37:E37").Formula = "=SUM(E38:E39)"
    xlHoja1.Range("E52:E52").Formula = "=+E12+E48+E46"
    
    xlHoja1.Range("F12:F12").Formula = "=+E12+D12"
    xlHoja1.Range("F17:F17").Formula = "=SUM(F18:F27)"
    xlHoja1.Range("F30:F30").Formula = "=SUM(F31:F33)"
    xlHoja1.Range("F37:F37").Formula = "=SUM(F38:F39)"
    xlHoja1.Range("F52:F52").Formula = "=+F12+F48+F46"
    
    xlHoja1.Range("G12:G12").Formula = "=+G14+G15+G16+G17+G28+G29+G30+G34+G35+G36+G37+G40+G41+G42+G43+G44"
    xlHoja1.Range("G17:G17").Formula = "=SUM(G18:G27)"
    xlHoja1.Range("G30:G30").Formula = "=SUM(G31:G33)"
    xlHoja1.Range("G37:G37").Formula = "=SUM(G38:G39)"
    xlHoja1.Range("G52:G52").Formula = "=+G12+G48+G46"
    
    xlHoja1.Range("H12:H12").Formula = "=+H14+H15+H16+H17+H28+H29+H30+H34+H35+H36+H37+H40+H41+H42+H43+H44"
    xlHoja1.Range("H17:H17").Formula = "=SUM(H18:H27)"
    xlHoja1.Range("H30:H30").Formula = "=SUM(H31:H33)"
    xlHoja1.Range("H37:H37").Formula = "=SUM(H38:H39)"
    xlHoja1.Range("H52:H52").Formula = "=+H12+H48+H46"
    
    xlHoja1.Range("I12:I12").Formula = "=+I14+I15+I16+I17+I28+I29+I30+I34+I35+I36+I37+I40+I41+I42+I43+I44"
    xlHoja1.Range("I17:I17").Formula = "=SUM(I18:I27)"
    xlHoja1.Range("I30:I30").Formula = "=SUM(I31:I33)"
    xlHoja1.Range("I37:I37").Formula = "=SUM(I38:I39)"
    xlHoja1.Range("I52:I52").Formula = "=+I12+I48+I46"
    
    xlHoja1.Range("A52:I52").Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(52, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:I52").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
    xlHoja1.Cells(53, 1) = "Periodicidad Mensual"
    xlHoja1.Cells(55, 1) = "(1) Clasificación industrial uniforme de todas las Actividades económicas. Tercera Revisión. Naciones Unidas. Consignar la actividad económica que genera el mayor valor añadido de la entidad deudora"
    xlHoja1.Cells(56, 1) = "(2) El total de créditos directos debe coincidir con la suma de las cuentas 1401+1403+1404+1405+1406+1407 del Manual de Contabilidad"
    xlHoja1.Cells(57, 1) = "(3) El total de créditos indirectos debe corresponder a la suma de los saldos de las cuentas 7101+7102+7103+7104+7105 del Manual de Contabilidad"
      
    xlHoja1.Range("A55:A57").Font.Bold = True
    xlHoja1.Range("A56:A60").Font.Size = 8
      
    xlHoja1.Range("D59:E59").MergeCells = True
    xlHoja1.Range("I59:J59").MergeCells = True
    xlHoja1.Range("O59:P59").MergeCells = True
     
    xlHoja1.Range("D60:E60").MergeCells = True
    xlHoja1.Range("I60:J60").MergeCells = True
    xlHoja1.Range("O60:P60").MergeCells = True
    
    xlHoja1.Range("I61:J61").MergeCells = True
      
    xlHoja1.Cells(60, 4) = "Gerente General"
    xlHoja1.Cells(60, 9) = "Contador General"
    xlHoja1.Cells(61, 9) = "Matricula Nro"
    xlHoja1.Cells(60, 15) = "Hecho Por"
     
    xlHoja1.Range("A60:T61").Font.Bold = True
    xlHoja1.Range("A60:T61").Font.Size = 8
     
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("B13:B61").HorizontalAlignment = xlCenter
     
    xlHoja1.Range("C11:C52").NumberFormat = "#,###,##0"
    xlHoja1.Range("G11:G52").NumberFormat = "#,###,##0"
   
    If gbBitCentral = True Then
        oConLocal.CierraConexion
    Else
        oCon.CierraConexion
    End If
   frmMdiMain.staMain.Panels(1).Text = ""
   RSClose R
End Sub

Private Sub CabeceraExcelAnexo3Riesgos(ByVal pdFecha As Date, psMes As String)
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = 46
    xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Cells(4, 1) = "EMPRESA: " & gsNomCmac
    xlHoja1.Cells(5, 9) = "Codigo: " & gsCodCMAC
    xlHoja1.Cells(3, 5) = "STOCK Y FLUJO CREDITICIO POR TIPO DE CREDITO Y SECTOR ECONOMICO"
    xlHoja1.Cells(4, 4) = "Al " & Mid(pdFecha, 1, 2) & " de " & Trim(psMes) & " de " & Year(pdFecha)
    xlHoja1.Cells(5, 4) = "( En Nuevos Soles )"
    xlHoja1.Cells(1, 10) = "ANEXO 3"
      
    xlHoja1.Range("C7:F7").Merge
    xlHoja1.Range("D8:F8").Merge
    xlHoja1.Range("G7:I7").Merge
    xlHoja1.Range("H8:I8").Merge
    
    xlHoja1.Range("A1:S9").HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A1:A50").ColumnWidth = 40
    xlHoja1.Range("B1:B50").ColumnWidth = 10
    xlHoja1.Range("C1:C50").ColumnWidth = 10
    xlHoja1.Range("D1:D50").ColumnWidth = 13
    xlHoja1.Range("E1:E50").ColumnWidth = 13
    xlHoja1.Range("F1:F50").ColumnWidth = 13
    xlHoja1.Range("G1:G50").ColumnWidth = 18
    xlHoja1.Range("H1:H50").ColumnWidth = 18
    xlHoja1.Range("I1:I50").ColumnWidth = 18
   
    xlHoja1.Range("B7:B10").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("A7:I10").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    xlHoja1.Range("A7:B10").Borders(xlInsideVertical).LineStyle = xlContinuous
    'xlHoja1.Range("C7:I10").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("C7:I7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D8:F8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("H8:I8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D7:I10").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:I10").HorizontalAlignment = xlHAlignCenter
    
    'xlHoja1.Cells(8, 1) = "Tipo de Credito"
    xlHoja1.Cells(8, 2) = "Division"
    xlHoja1.Cells(9, 2) = "CIIU (*)"
    xlHoja1.Cells(7, 3) = "STOCK AL CIERRE DEL MES"
    xlHoja1.Cells(8, 3) = "Numero"
    xlHoja1.Cells(9, 3) = "de"
    xlHoja1.Cells(10, 3) = "Deudores"
    xlHoja1.Cells(8, 4) = "Saldo"
    xlHoja1.Cells(9, 4) = "M.N."
    xlHoja1.Cells(10, 4) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 5) = "M.E."
    xlHoja1.Cells(10, 5) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 6) = "Total"
    xlHoja1.Cells(10, 6) = "(Miles de N.S.)"
   
    
    xlHoja1.Cells(7, 7) = "FLUJO DESEMBOLSADO EN EL MES"
    xlHoja1.Cells(8, 7) = "Número de nuevos"
    xlHoja1.Cells(9, 7) = "créditos"
    xlHoja1.Cells(10, 7) = "desembolsados"
    xlHoja1.Cells(8, 8) = "Monto de nuevos créditos desemb."
    xlHoja1.Cells(9, 8) = "M.N."
    xlHoja1.Cells(10, 8) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 9) = "M.E."
    xlHoja1.Cells(10, 9) = "(Miles de US$)"
    
    xlHoja1.Cells(9, 10) = " % "
    xlHoja1.Cells(10, 10) = "Mora"
    
    xlHoja1.Range("D12:F50").NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range("H12:I50").NumberFormat = "#,##0.00;-#,##0.00"
End Sub

'******************************************************************************



