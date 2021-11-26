VERSION 5.00
Begin VB.Form frmRepSpreadFinanciero 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte 06: Reporte Diario de Tasas de Interés"
   ClientHeight    =   690
   ClientLeft      =   1350
   ClientTop       =   2340
   ClientWidth     =   5565
   Icon            =   "frmRepSpreadFinanciero.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmRepSpreadFinanciero"
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

Dim ldFecha  As Date
Dim oBarra As clsProgressBar
Dim oCon As DConecta
'******JUCS****2017/02/08********************
Dim nGuardaMN As Double
Dim nGuardaME As Double
Dim nGuardaTotMN As Double
Dim nGuardaTotME As Double
Dim TpoCambio As Double
Dim nTaeMN As Double
Dim nSaldoMN As Double
Dim nSubTotalMN As Double
Dim nSubTotalME As Double
Dim nTotalMN As Double
Dim nTotalME As Double
Dim nConsCell As Double
Dim nAlmacTsa As Double
Dim nTotSaldoMN As Double
Dim nTotSaldoME As Double
'Calculo Spread
Dim SumTotSpread As Double
Dim SumTotRep6A As Double
Dim SpreadMN As Double
Dim SpreadME As Long
Dim TotSpread As Long
Dim nTotSaldoMNX As Double
'variables finales Spread
Dim ResultSpread() As Variant
Dim TAMN As Double
Dim TAME As Double
Dim TAO As Double
'******JUCS****2017/02/08*******************
Dim sservidorconsolidada As String
Public Sub ImprimeSpreadFinanciero(psOpeCod As String, pdFecha As Date, ByVal sLstage As ListBox)
On Error GoTo GeneraEstadError
Dim oConecta As DConecta
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim cConsol As String
ReDim ResultSpread(8)

    Dim I As Integer
    Dim lsAgencia As String
    
    lsAgencia = ""
    cConsol = "N"
    For I = 0 To sLstage.ListCount - 1
        If sLstage.Selected(I) Then
            If Right(sLstage.List(I), 5) = "ONSOL" Then
                lsAgencia = lsAgencia & Right(sLstage.List(I), 2) & ","
                cConsol = "S"
            Else
                lsAgencia = lsAgencia & Right(sLstage.List(I), 2) & ","
            End If
        End If
    Next I
    
    lsAgencia = Mid(lsAgencia, 1, Len(lsAgencia) - 1)


   Set oConecta = New DConecta
   oConecta.AbreConexion
   Set rs = oConecta.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
    If rs.BOF Then
    Else
        sservidorconsolidada = rs!nConsSisValor
    End If
    Set rs = Nothing
   
   oConecta.CierraConexion
   Set oConecta = Nothing
   
   Set oCon = New DConecta
   oCon.AbreConexion 'Remota gsCodAge, True, False, "03"
   
   ldFecha = pdFecha
   lsArchivo = App.path & "\SPOOLER\" & "SBSRepSpreadFinanc_" & Format(pdFecha, "mmyyyy") & ".XLSX"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbExcel Then
   
      ExcelAddHoja "Rep_6A", xlLibro, xlHoja1
      Genera6A 1, pdFecha, lsAgencia, cConsol 'Coloc Ok
      'Genera6A 2, pdFecha, lsAgencia, cConsol 'Coloc OK
      Genera6A_N 2, pdFecha, lsAgencia, cConsol 'Coloc OK

      ExcelAddHoja "Rep_6B", xlLibro, xlHoja1
      Genera6B 1, pdFecha, lsAgencia, cConsol 'Capta OK
      Genera6B 2, pdFecha, lsAgencia, cConsol 'Capta OK

      ExcelAddHoja "Rep_6D (nuevo)", xlLibro, xlHoja1
      Genera6D1 1, pdFecha, lsAgencia, cConsol 'Coloc oK
      Genera6D1 2, pdFecha, lsAgencia, cConsol 'Coloc Ok

      ExcelAddHoja "Rep_6E (nuevo)", xlLibro, xlHoja1
      Genera6E1 1, pdFecha, lsAgencia, cConsol 'Capta
      Genera6E1 2, pdFecha, lsAgencia, cConsol 'Capta
      
      ExcelAddHoja "Spread", xlLibro, xlHoja1 'JUCS 15032017
      GeneraSpread 1, pdFecha, lsAgencia, cConsol

      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      If lsArchivo <> "" Then
         CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      End If
   End If
   oCon.CierraConexion
   Set oCon = Nothing
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub Genera6A(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim sSqlTemp As String
    Dim rs As New ADODB.Recordset
    Dim cPigno As String
    Dim cVigente As String
    
    'sSql = "SELECT nValComp " 'JUCS 20170203
    ' sSql = sSql & "FROM TipoCambio "
    'sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,('" & Format(pdFecha, gsFormatoFecha) & "')) =0 "
     
     sSql = "SELECT nValFijo " 'JUCS 20170203
     sSql = sSql & "FROM TipoCambio "
     sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,dateadd(day,1,('" & Format(pdFecha, gsFormatoFecha) & "'))) =0"
    
    Set rs = oCon.CargaRecordSet(sSql)
    xlHoja1.Cells(lnFila + 5, 7) = "T/C"
    xlHoja1.Cells(lnFila + 5, 8) = rs!nValFijo
    TpoCambio = rs!nValFijo
    
    'GCOLOCESTRECREFJUD
    'xxx cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "'"
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "' , '" & gColocEstRecVigJud & "','2205'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
   
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS SOBRE SALDOS", "REPORTE 6A", 8)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 7) = "CALCULO SPREAD" 'JUCS 20170203
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    xlHoja1.Cells(lnFila + 1, 7) = "MN" 'JUCS 20170203
    xlHoja1.Cells(lnFila + 1, 8) = "ME" 'JUCS 20170203
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    xlHoja1.Range("G1").ColumnWidth = 17 'JUCS 20170203
    xlHoja1.Range("H1").ColumnWidth = 17 'JUCS 20170203
    xlHoja1.Range("I1").ColumnWidth = 17 'JUCS 20170203
    xlHoja1.Range("J1").ColumnWidth = 17 'JUCS 20170203
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).MergeCells = True 'JUCS 20170203
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2

        sSql = " SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        'sSql = sSql & " ISNULL(round(SUM(nTasaInt)/count(*),2),0) nTasaInt_Mes, "
        sSql = sSql & " round(ISNULL(SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "  THEN case "
        sSql = sSql & "   when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "  then  0 "
        sSql = sSql & "   else  nTasaInt END "
        sSql = sSql & "  else nTasaInt end)/ "
        sSql = sSql & "  SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "     THEN case"
        sSql = sSql & "       when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "                        then  0 "
        sSql = sSql & "    else  1 end "
        sSql = sSql & "      else 1 end),0),2) nTasaInt_Mes, "
        'sSql = sSql & " ISNULL((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol),0) nTasaInt_Anio, "
        sSql = sSql & "  ROUND(ISNULL((SUM( "
        sSql = sSql & "      CASE WHEN cSitCtb='2' "
        sSql = sSql & "             THEN CASE "
        sSql = sSql & "                 WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                     THEN  0"
        sSql = sSql & "                 ELSE  (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "                 End "
        sSql = sSql & "            ELSE (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "      End "
        sSql = sSql & "  ))/SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "              THEN CASE "
        sSql = sSql & "                        WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                          THEN 0"
        sSql = sSql & "              ELSE  col.nMontoCol END"
        sSql = sSql & "     ELSE "
        sSql = sSql & "             Col.nMontoCol "
        sSql = sSql & "     end),0),2) nTasaInt_Anio,"
        'sSql = sSql & " ISNULL(SUM(CASE WHEN C.cTpoCredCod LIKE '7%' AND NOT C.cTpoCredCod = '755' and c.nDiasAtraso >=31 and c.nDiasAtraso <=90 THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " ISNULL(SUM(CASE WHEN cSitCtb='2' THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " FROM  " & sservidorconsolidada & "Rangox r "
        sSql = sSql & " LEFT JOIN( "
        sSql = sSql & " SELECT cc.cSitCtb,c.cCtaCod, Ct.cTpoCredCod, ct.nTasaInt, c.nDiasAtraso, c.nSaldoCap, ct.nCuotasApr, c.nCapVencido, "
        sSql = sSql & " ct.nCuotasApr * CASE WHEN ct.nPlazoApr = 0 THEN 30 ELSE ct.nPlazoApr END nDias "
        sSql = sSql & " FROM " & sservidorconsolidada & "CreditoSaldoConsol c "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsolTotal ct ON ct.cCtaCod = c.cCtaCod "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsol cc ON ct.cCtaCod = cc.cCtaCod "
        sSql = sSql & " WHERE  DATEDIFF(D,DFECHA,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And ct.cRefinan = 'N' "
        sSql = sSql & " AND c.nPrdEstado IN (" & cVigente & ", " & cPigno & ") "
        sSql = sSql & " AND c.nSaldoCap > 0 AND SUBSTRING(c.cCtaCod,9,1) = " & psMoneda
        'sSql = sSql & " AND NOT Ct.cTpoCredCod IN ('756') "
        'sSql = sSql & "     AND ("
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[123]%' AND c.nDiasAtraso < 16) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[45]%' AND c.nDiasAtraso < 31) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[678]%' AND NOT ct.cTpoCredCod = '755' AND c.nDiasAtraso <= 90) or"
        'sSql = sSql & "       (CT.cTpoCredCod='755' AND c.nDiasAtraso < 31)"
        'sSql = sSql & "     )"
        sSql = sSql & " and cc.cSitCtb in ('1','2')"
        If psConsol = "N" Then
            sSql = sSql & " And CT.cAgeCodAct in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql & " ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin And nSaldoCap"
        sSql = sSql & "   Between R.nMontoIni And R.nMontoFin And R.cProdCod = SubString(c.cTpoCredCod, 1, 1)"
'        sSql = sSql & "   AND ((r.bPrdIn = 1  And C.cTpoCredCod"
'        sSql = sSql & "   IN (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
'        sSql = sSql & "     OR (r.bPrdOut = 1 And C.cTpoCredCod"
'        sSql = sSql & "     NOT In (SELECT cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
        sSql = sSql & "   LEFT JOIN Colocaciones col on (C.cCtaCod=col.cCtaCod)"
        sSql = sSql & " WHERE r.cTipoAnx = 'A' AND r.nMoneda = " & psMoneda & "  and r.nRangoCod not in (36)   GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod"
        sSql = sSql & " ORDER BY nRangoCod"
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
     
    Dim Pro1, Pro2, Pro1X, Pro2X As Double
    Dim SumSaldo, SumSaldoX, SumSaldoM As Double
    nTotalMN = 0
    nTotSaldoMN = 0
    nTotalME = 0
    nTotSaldoME = 0
   
    Do While Not rs.EOF
     If psMoneda = "1" Then
            'aqui llena la primera columna
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            'si producto es diferente de 1 o del record set entonces dar formato de bordes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        'aqui comienza el llenado de datos numericos
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
            If lnFila = 39 Or lnFila = 40 Then
                SumSaldoM = rs!nSaldo
              
                'If lnFila = 39 Then
                rs.MoveNext
                'End If
                If lnFila = 39 Then
                Pro1 = rs!nTasaInt_Anio * rs!nSaldo
                SumSaldo = rs!nSaldo
                End If
                If lnFila = 40 Then
                Pro1X = rs!nTasaInt_Anio * rs!nSaldo 'vapa
                SumSaldoX = rs!nSaldo 'vapa
                End If
                'If lnFila = 39 Then
                rs.MoveNext
                'End If
                If lnFila = 39 Then
                Pro2 = rs!nTasaInt_Anio * rs!nSaldo
                SumSaldo = SumSaldo + rs!nSaldo
                End If
                If lnFila = 40 Then
                Pro2X = rs!nTasaInt_Anio * rs!nSaldo
                SumSaldoX = SumSaldoX + rs!nSaldo                'vapa
                End If
                
                If lnFila = 39 Then
                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = (Pro1 + Pro2) / SumSaldo
                Else
                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = (Pro1X + Pro2X) / SumSaldoX
                End If
                            If lnFila = 39 Then
                                nAlmacTsa = (Pro1 + Pro2) / SumSaldo
                                nConsCell = SumSaldo * nAlmacTsa
                                xlHoja1.Cells(lnFila, 7) = nConsCell
                            End If
                            If lnFila = 40 Then
                                nAlmacTsa = (Pro1X + Pro2X) / SumSaldoX
                                nConsCell = SumSaldoM * nAlmacTsa
                                xlHoja1.Cells(lnFila, 7) = nConsCell
                                
                            End If
            Else
                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            End If
                  '******JUCS****2017/02/08**********************************************************************************
                If psMoneda = 1 Then
                     nTaeMN = rs!nTasaInt_Anio
                     nSaldoMN = rs!nSaldo
                     nSubTotalMN = nSaldoMN * nTaeMN
'                     If lnFila = 39 Or lnFila = 40 Then
'                         nAlmacTsa = (Pro1 + Pro2) / SumSaldo.
'                         nConsCell = SumSaldo * nAlmacTsa
'                          xlHoja1.Cells(lnFila, 7) = nConsCell
'                     Else
                          If Not lnFila = 39 And Not lnFila = 40 Then
                            xlHoja1.Cells(lnFila, 7) = nSubTotalMN
                          End If
'                     End If
                      'Almacenan Total de saldo en MN
                    If lnFila = 39 Or lnFila = 40 Then
                        nTotalMN = nTotalMN + nConsCell
                        If lnFila = 40 Then
                           nTotSaldoMN = nTotSaldoMN + SumSaldoM
                        Else
                           nTotSaldoMN = nTotSaldoMN + SumSaldo
                      End If
                    Else
                        nTotalMN = nTotalMN + nSubTotalMN
                         nTotSaldoMN = nTotSaldoMN + nSaldoMN
                    End If
                Else
                     nTaeMN = rs!nTasaInt_Anio
                     nSaldoMN = rs!nSaldo
                     nSubTotalME = nSaldoMN * nTaeMN
                     'Almacenan Total de saldo en ME
                     nTotalME = nTotalME + nSubTotalME
                     nTotSaldoME = nTotSaldoME + nSaldoMN
                     xlHoja1.Cells(lnFila, 8) = nSubTotalME
                End If
                  '******JUCS****2017/02/08**********************************************************************************
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    '******JUCS****2017/02/08**********************************************************************************
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium 'jucs 20170202
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium 'jucs 20170202
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous 'jucs 20170202
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous 'jucs 20170202
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00" 'jucs 20170202
    End If
    'Totales
    If psMoneda = "1" Then
       xlHoja1.Cells(lnFila + 1, 1) = "TOTAL :"
       xlHoja1.Cells(lnFila + 1, 7) = nTotalMN
       xlHoja1.Cells(lnFila + 1, 4) = nTotSaldoMN
    Else
       xlHoja1.Cells(lnFila + 1, 8) = nTotalME
       xlHoja1.Cells(lnFila + 1, 6) = nTotSaldoME
    End If
    'Calclulo Final y llenado en celdas de Spread Financiero
    If psMoneda = "1" Then
       nGuardaMN = nTotalMN
       nGuardaTotMN = nTotSaldoMN
       
       If nGuardaTotMN = 0 Then
       TAMN = 0
       Else
       TAMN = nGuardaMN / nGuardaTotMN
       End If
       
       ResultSpread(0) = TAMN
       ReDim Preserve ResultSpread(8)
    Else
        nGuardaMN = nGuardaMN
        nGuardaTotMN = nGuardaTotMN
        SumTotSpread = nGuardaMN + nTotalME * TpoCambio
        SumTotRep6A = nGuardaTotMN + nTotSaldoME * TpoCambio
        If nTotSaldoME = 0 Then
        TAME = 0
        Else
        TAME = nTotalME / nTotSaldoME
        End If
        If SumTotRep6A = 0 Then
        TAO = 0
        Else
        TAO = SumTotSpread / SumTotRep6A
        End If
        ResultSpread(1) = TAME
        ReDim Preserve ResultSpread(8)
        ResultSpread(2) = TAO
        ReDim Preserve ResultSpread(8)
        
        xlHoja1.Cells(lnFila + 1, 9) = SumTotSpread
        xlHoja1.Cells(lnFila + 1, 10) = SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 1) = "SPREAD FINANCIERO"
        xlHoja1.Cells(lnFila + 2, 7) = TAMN
        xlHoja1.Cells(lnFila + 2, 8) = TAME
        xlHoja1.Cells(lnFila + 2, 9) = TAO
     End If
        xlHoja1.Range("A44:J44").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A45:J45").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A46:J46").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A44:J44").NumberFormat = "#,##0.00"
        xlHoja1.Range("A45:J45").NumberFormat = "#,##0.00"
     '****** FIN JUCS****2017/02/08**********************************************************************************
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
'VAPA 022017
Private Sub Genera6A_N(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim sSqlTemp As String
    Dim rs As New ADODB.Recordset
    Dim cPigno As String
    Dim cVigente As String
    
    'sSql = "SELECT nValComp " 'JUCS 20170203
    'sSql = sSql & "FROM TipoCambio "
    'sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,('" & Format(pdFecha, gsFormatoFecha) & "')) =0 "
    
     sSql = "SELECT nValFijo " 'JUCS 20170203
     sSql = sSql & "FROM TipoCambio "
     sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,dateadd(day,1,('" & Format(pdFecha, gsFormatoFecha) & "'))) =0"
    
    Set rs = oCon.CargaRecordSet(sSql)
    xlHoja1.Cells(lnFila + 5, 7) = "T/C"
    xlHoja1.Cells(lnFila + 5, 8) = rs!nValFijo
    TpoCambio = rs!nValFijo
    
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "' , '" & gColocEstRecVigJud & "','2205'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
   
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS SOBRE SALDOS", "REPORTE 6A", 8)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 7) = "CALCULO SPREAD" 'JUCS 20170203
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    xlHoja1.Cells(lnFila + 1, 7) = "MN" 'JUCS 20170203
    xlHoja1.Cells(lnFila + 1, 8) = "ME" 'JUCS 20170203
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    xlHoja1.Range("G1").ColumnWidth = 18 'JUCS 20170203
    xlHoja1.Range("H1").ColumnWidth = 18 'JUCS 20170203
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).MergeCells = True 'JUCS 20170203
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
        sSql = " SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        'sSql = sSql & " ISNULL(round(SUM(nTasaInt)/count(*),2),0) nTasaInt_Mes, "
        sSql = sSql & " round(ISNULL(SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "  THEN case "
        sSql = sSql & "   when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "  then  0 "
        sSql = sSql & "   else  nTasaInt END "
        sSql = sSql & "  else nTasaInt end)/ "
        sSql = sSql & "  SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "     THEN case"
        sSql = sSql & "       when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "                        then  0 "
        sSql = sSql & "    else  1 end "
        sSql = sSql & "      else 1 end),0),2) nTasaInt_Mes, "
        'sSql = sSql & " ISNULL((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol),0) nTasaInt_Anio, "
        sSql = sSql & "  ROUND(ISNULL((SUM( "
        sSql = sSql & "      CASE WHEN cSitCtb='2' "
        sSql = sSql & "             THEN CASE "
        sSql = sSql & "                 WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                     THEN  0"
        sSql = sSql & "                 ELSE  (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "                 End "
        sSql = sSql & "            ELSE (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "      End "
        sSql = sSql & "  ))/SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "              THEN CASE "
        sSql = sSql & "                        WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                          THEN 0"
        sSql = sSql & "              ELSE  col.nMontoCol END"
        sSql = sSql & "     ELSE "
        sSql = sSql & "             Col.nMontoCol "
        sSql = sSql & "     end),0),2) nTasaInt_Anio,"
        'sSql = sSql & " ISNULL(SUM(CASE WHEN C.cTpoCredCod LIKE '7%' AND NOT C.cTpoCredCod = '755' and c.nDiasAtraso >=31 and c.nDiasAtraso <=90 THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " ISNULL(SUM(CASE WHEN cSitCtb='2' THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " FROM  " & sservidorconsolidada & "Rangox r "
        sSql = sSql & " LEFT JOIN( "
        sSql = sSql & " SELECT cc.cSitCtb,c.cCtaCod, Ct.cTpoCredCod, ct.nTasaInt, c.nDiasAtraso, c.nSaldoCap, ct.nCuotasApr, c.nCapVencido, "
        sSql = sSql & " ct.nCuotasApr * CASE WHEN ct.nPlazoApr = 0 THEN 30 ELSE ct.nPlazoApr END nDias "
        sSql = sSql & " FROM " & sservidorconsolidada & "CreditoSaldoConsol c "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsolTotal ct ON ct.cCtaCod = c.cCtaCod "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsol cc ON ct.cCtaCod = cc.cCtaCod "
        sSql = sSql & " WHERE  DATEDIFF(D,DFECHA,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And ct.cRefinan = 'N' "
        sSql = sSql & " AND c.nPrdEstado IN (" & cVigente & ", " & cPigno & ") "
        sSql = sSql & " AND c.nSaldoCap > 0 AND SUBSTRING(c.cCtaCod,9,1) = " & psMoneda
        'sSql = sSql & " AND NOT Ct.cTpoCredCod IN ('756') "
        'sSql = sSql & "     AND ("
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[123]%' AND c.nDiasAtraso < 16) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[45]%' AND c.nDiasAtraso < 31) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[678]%' AND NOT ct.cTpoCredCod = '755' AND c.nDiasAtraso <= 90) or"
        'sSql = sSql & "       (CT.cTpoCredCod='755' AND c.nDiasAtraso < 31)"
        'sSql = sSql & "     )"
        sSql = sSql & " and cc.cSitCtb in ('1','2')"
        If psConsol = "N" Then
            sSql = sSql & " And CT.cAgeCodAct in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql & " ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin And nSaldoCap"
        sSql = sSql & "   Between R.nMontoIni And R.nMontoFin And R.cProdCod = SubString(c.cTpoCredCod, 1, 1)"
'        sSql = sSql & "   AND ((r.bPrdIn = 1  And C.cTpoCredCod"
'        sSql = sSql & "   IN (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
'        sSql = sSql & "     OR (r.bPrdOut = 1 And C.cTpoCredCod"
'        sSql = sSql & "     NOT In (SELECT cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
        sSql = sSql & "   LEFT JOIN Colocaciones col on (C.cCtaCod=col.cCtaCod)"
        sSql = sSql & " WHERE r.cTipoAnx = 'A' AND r.nMoneda = " & psMoneda & " and r.nRangoCod not in (35,31,32,34,36)  GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod"
        sSql = sSql & " ORDER BY nRangoCod"
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
     
    Dim Pro1, Pro2 As Double
    Dim SumSaldo As Double
   nTotalMN = 0
   nTotSaldoMN = 0
   nTotalME = 0
   nTotSaldoME = 0
   
    Do While Not rs.EOF
        If psMoneda = "1" Then
            'aqui llena la primera columna
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            'si producto es diferente de 1 o del record set entonces dar formato de bordes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        'aqui comienza el llenado de datos numericos
        If rs!nSaldo <> 0 Then
                xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo

                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            'End If
                  '******JUCS****2017/02/08**********************************************************************************
                  If psMoneda = 1 Then
                     nTaeMN = rs!nTasaInt_Anio
                     nSaldoMN = rs!nSaldo
                     nSubTotalMN = nSaldoMN * nTaeMN
                     
                     If lnFila = 39 Or lnFila = 40 Then
                         nAlmacTsa = (Pro1 + Pro2) / SumSaldo
                         nConsCell = SumSaldo * nAlmacTsa
                          xlHoja1.Cells(lnFila, 7) = nConsCell
                     Else
                          xlHoja1.Cells(lnFila, 7) = nSubTotalMN
                     End If
                      'Almacenan Total de saldo en MN
                     If lnFila = 39 Or lnFila = 40 Then
                        nTotalMN = nTotalMN + nConsCell
                        nTotSaldoMN = nTotSaldoMN + SumSaldo
                     Else
                        nTotalMN = nTotalMN + nSubTotalMN
                         nTotSaldoMN = nTotSaldoMN + nSaldoMN
                     End If
                  Else
                     nTaeMN = rs!nTasaInt_Anio
                     nSaldoMN = rs!nSaldo
                     nSubTotalME = nSaldoMN * nTaeMN
                     'Almacenan Total de saldo en ME
                     nTotalME = nTotalME + nSubTotalME
                     nTotSaldoME = nTotSaldoME + nSaldoMN
                     xlHoja1.Cells(lnFila, 8) = nSubTotalME
                  End If
                  '******JUCS****2017/02/08**********************************************************************************
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop

    'Totales
    If psMoneda = "1" Then
       xlHoja1.Cells(lnFila + 1, 7) = nTotalMN
       xlHoja1.Cells(lnFila + 1, 4) = nTotSaldoMN
      
    Else
       xlHoja1.Cells(lnFila + 1, 8) = nTotalME
       xlHoja1.Cells(lnFila + 1, 6) = nTotSaldoME
    End If  'jucs 20170202
       
    'Calclulo Final y llenado en celdas de Spread Financiero
    If psMoneda = "1" Then
       nGuardaMN = nTotalMN
       nGuardaTotMN = nTotSaldoMN
    Else
        nGuardaMN = nGuardaMN
        nGuardaTotMN = nGuardaTotMN
        SumTotSpread = nGuardaMN + nTotalME * TpoCambio
        SumTotRep6A = nGuardaTotMN + nTotSaldoME * TpoCambio
        If nTotSaldoME = 0 Then
        TAME = 0
        Else
        TAME = nTotalME / nTotSaldoME
        End If
        If SumTotRep6A = 0 Then
        TAO = 0
        Else
        TAO = SumTotSpread / SumTotRep6A
        End If
      
       ResultSpread(1) = TAME
       ReDim Preserve ResultSpread(8)
       
      
       ResultSpread(2) = TAO
       ReDim Preserve ResultSpread(8)
        
        xlHoja1.Cells(lnFila + 1, 9) = SumTotSpread
        xlHoja1.Cells(lnFila + 1, 10) = SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 1) = "SPREAD FINANCIERO"
'        xlHoja1.Cells(lnFila + 2, 7) = nGuardaMN / nGuardaTotMN
'        xlHoja1.Cells(lnFila + 2, 8) = nTotalME / nTotSaldoME
'        xlHoja1.Cells(lnFila + 2, 9) = SumTotSpread / SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 7) = TAMN
        xlHoja1.Cells(lnFila + 2, 8) = TAME
        xlHoja1.Cells(lnFila + 2, 9) = TAO
     End If
     '****** FIN JUCS****2017/02/08**********************************************************************************
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
'end VAPA

Private Function CabeceraReporte(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer) As Integer
Dim lnFila As Integer
    xlHoja1.Range("A1:R100").Font.Size = 8

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
    
    lnFila = 1
    xlHoja1.Cells(lnFila, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 2) = psTitulo
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 5) = psReporte
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 1) = "EMPRESA : " & gsNomCmac:
    xlHoja1.Cells(lnFila + 1, 1) = "Fecha : AL " & Format(pdFecha, "dd mmmm yyyy")
    
    lnFila = lnFila + 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Font.Bold = True
    CabeceraReporte = lnFila
End Function

Private Sub Genera6B(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
    'sSql = "SELECT nValComp " 'JUCS
    'sSql = sSql & "FROM TipoCambio "
    'sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,('" & Format(pdFecha, gsFormatoFecha) & "')) =0 "
    
     sSql = "SELECT nValFijo " 'JUCS 20170203
     sSql = sSql & "FROM TipoCambio "
     sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,dateadd(day,1,('" & Format(pdFecha, gsFormatoFecha) & "'))) =0"
     
    Set rs = oCon.CargaRecordSet(sSql)
    xlHoja1.Cells(lnFila + 5, 7) = "T/C"
    xlHoja1.Cells(lnFila + 5, 8) = rs!nValFijo
    TpoCambio = rs!nValFijo 'JUCS
    
If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS SOBRE SALDOS", "REPORTE 6B", 8)
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    xlHoja1.Range("G1").ColumnWidth = 18 'JUCS 20170208
    xlHoja1.Range("H1").ColumnWidth = 18 'JUCS 20170208
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 7) = "CALCULO SPREAD" 'JUCS 20170208
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
     xlHoja1.Cells(lnFila + 1, 7) = "MN" 'JUCS 20170208
    xlHoja1.Cells(lnFila + 1, 8) = "ME" 'JUCS 20170208
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).MergeCells = True 'JUCS 20170208
Else
    lnFila = 10
End If

    oBarra.Progress 1, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    If gbBitCentral = True Then
        sSql = "         SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        sSql = sSql + "         ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, "
        sSql = sSql + "         ISNULL(SUM(nSaldo),0) nSaldo"
        sSql = sSql + "  FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "           LEFT JOIN     ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual,"
        'sSql = sSql + "                                 ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldCntCTS)/sum(nSaldCntCTS) nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntCTS) nSaldo"
        sSql = sSql + "                            From " & sservidorconsolidada & "CTSConsol "
        'sSql = sSql + "                            inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "CTSConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                            where   nEstCtaCTS not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "'  "
        If psConsol = "N" Then
            sSql = sSql & "                                     And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                             Union"
        sSql = sSql + "                             Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                     ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                 SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/sum(nSaldCntAc) nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntAc) nSaldo"
        sSql = sSql + "                             From " & sservidorconsolidada & "AhorroCConsol"
        'sSql = sSql + "                             inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "AhorroCConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             where   nEstCtaAC not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria in (1,2,3) And bInactiva = 0 "
        If psConsol = "N" Then
            sSql = sSql & "                                     And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                             Union"
        sSql = sSql + "                             SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                             ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio,"
        sSql = sSql + "                                     SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                             FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                     JOIN " & sservidorconsolidada & "Rango R1 ON pf.nPlazo "
        sSql = sSql + "                                     Between R1.nRangoIni And R1.nRangoFin "
        'sSql = sSql + "                                     inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "PlazoFijoConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria In (1,2,3) And cTipoAnx = 'B' "
        sSql = sSql + "                                     AND NOT EXISTS (SELECT  PC.CCTACOD "
        sSql = sSql + "                                                     FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                     WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                             cMovNroDbl IS NULL "
        sSql = sSql + "                                                     AND nBlqMotivo = 3)"
        If psConsol = "N" Then
            sSql = sSql & "                                     AND SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                                     GROUP BY R1.nRangoIni) Dat "
        sSql = sSql + "                                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql + "                                     WHERE cTipoAnx = 'B' AND nRangoCod <> 20"
        sSql = sSql + "                              GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql + "      Union"
        sSql = sSql + "      SELECT Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
        sSql = sSql + "             SUM(nSaldo) As nSaldo "
        sSql = sSql + "      From"
        sSql = sSql + "           (SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes,"
        sSql = sSql + "                   ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
        sSql = sSql + "                    ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
        sSql = sSql + "            FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "                    JOIN     (  "
        sSql = sSql + "                                 SELECT  -1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio, "
        sSql = sSql + "                                         SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                                  FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                  WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' "
        sSql = sSql + "                                          AND EXISTS (   SELECT  PC.CCTACOD "
        sSql = sSql + "                                                         FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                         WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                                 PF.nEstCtaPF IN (1100,1200) AND cMovNroDbl IS NULL "
        sSql = sSql + "                                                                 AND nBlqMotivo = 3)"
        If psConsol = "N" Then
            sSql = sSql & "                                          And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                                             ) Dat "
        sSql = sSql + "                                 ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
        sSql = sSql + "                                 WHERE cTipoAnx = 'B' AND nRangoCod = 20 "
        sSql = sSql + "                                 GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio ) AS x "
        sSql = sSql + "                                 GROUP BY Prod, nRangoCod, nRangoDes"
        sSql = sSql + "                                 ORDER BY Prod, nRangoCod, nRangoDes" 'EJVG20131202
        '*********************************************************************************************************************************************************
        
    Else
        sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, ISNULL(SUM(nSaldo),0) nSaldo "
        sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
        sSql = sSql & "    ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntCTS) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "CTSConsol where cEstCtaCTS not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/sum(nSaldCntAc) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        'sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        sSql = sSql & "      From " & sservidorconsolidada & "AhorroCConsol "
        sSql = sSql & "      where cEstCtaAC not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        'sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "PlazoFijoConsol pf JOIN Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
        sSql = sSql & "      WHERE cEstCtaPF not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' and cTipoAnx = 'B' "
        sSql = sSql & "      GROUP BY R1.nRangoIni "
        sSql = sSql & "    ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & "WHERE cTipoAnx = 'B' "
        sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql & "ORDER BY r.nRangoCod"
    End If
    '*********************************************************************************************************************************************************
    '**END***2008/06/03**************************************************************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
   nTotalMN = 0
   nTotSaldoMN = 0
   nTotalME = 0
   nTotSaldoME = 0
     
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
             '******JUCS****2017/02/08**********************************************************************************
            If psMoneda = 1 Then
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalMN = nSaldoMN * nTaeMN
            nTotalMN = nTotalMN + nSubTotalMN
            nTotSaldoMN = nTotSaldoMN + nSaldoMN
            xlHoja1.Cells(lnFila, 7) = nSubTotalMN
            Else
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalME = nSaldoMN * nTaeMN
            nTotalME = nTotalME + nSubTotalME
            nTotSaldoME = nTotSaldoME + nSaldoMN
            xlHoja1.Cells(lnFila, 8) = nSubTotalME
            End If
            '******FIN JUCS****2017/02/08********************************************************************************
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    '******JUCS****2017/02/08**********************************************************************************
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00"
    End If
    If psMoneda = "1" Then 'jucs 20170208
       xlHoja1.Cells(lnFila + 1, 1) = "TOTAL :"
       xlHoja1.Cells(lnFila + 1, 7) = nTotalMN
       xlHoja1.Cells(lnFila + 1, 4) = nTotSaldoMN
    Else
       xlHoja1.Cells(lnFila + 1, 8) = nTotalME
       xlHoja1.Cells(lnFila + 1, 6) = nTotSaldoME
    End If  'jucs 20170202
    
     'Calclulo totales
    If psMoneda = "1" Then
       nGuardaMN = nTotalMN
       nGuardaTotMN = nTotSaldoMN
       If nGuardaTotMN = 0 Then
       TAMN = 0
       Else
       TAMN = nGuardaMN / nGuardaTotMN
       End If
      
       ResultSpread(3) = TAMN
       ReDim Preserve ResultSpread(8)
    Else
        nGuardaMN = nGuardaMN
        nGuardaTotMN = nGuardaTotMN
        SumTotSpread = nGuardaMN + nTotalME * TpoCambio
        SumTotRep6A = nGuardaTotMN + nTotSaldoME * TpoCambio
        If nTotSaldoME = 0 Then
        TAME = 0
        Else
        TAME = nTotalME / nTotSaldoME
        End If
        If SumTotRep6A = 0 Then
        TAO = 0
        Else
        TAO = SumTotSpread / SumTotRep6A
        End If
       
        ResultSpread(4) = TAME
        ReDim Preserve ResultSpread(8)
        
        ResultSpread(5) = TAO
        ReDim Preserve ResultSpread(8)
        
        xlHoja1.Cells(lnFila + 1, 9) = SumTotSpread
        xlHoja1.Cells(lnFila + 1, 10) = SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 1) = "SPREAD FINANCIERO"
'       xlHoja1.Cells(lnFila + 2, 7) = nGuardaMN / nGuardaTotMN
'       xlHoja1.Cells(lnFila + 2, 8) = nTotalME / nTotSaldoME
'       xlHoja1.Cells(lnFila + 2, 9) = SumTotSpread / SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 7) = TAMN
        xlHoja1.Cells(lnFila + 2, 8) = TAME
        xlHoja1.Cells(lnFila + 2, 9) = TAO
     End If
        xlHoja1.Range("A38:J39").Font.Bold = True
        xlHoja1.Range("A38:J38").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A39:J39").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A40:J40").Borders(xlEdgeTop).LineStyle = True
        xlHoja1.Range("A38:J38").NumberFormat = "#,##0.00"
        xlHoja1.Range("A39:J39").NumberFormat = "#,##0.00"
    '****** FIN JUCS****2017/02/08**********************************************************************************
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
''peac 20071127

Private Sub Genera6D1(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim cJudicial As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    'sSql = "SELECT nValComp " 'JUCS
    'sSql = sSql & "FROM TipoCambio "
    'sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,('" & Format(pdFecha, gsFormatoFecha) & "')) =0 "
    
     sSql = "SELECT nValFijo " 'JUCS 20170203
     sSql = sSql & "FROM TipoCambio "
     sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,dateadd(day,1,('" & Format(pdFecha, gsFormatoFecha) & "'))) =0"
    
    Set rs = oCon.CargaRecordSet(sSql)
    xlHoja1.Cells(lnFila + 5, 9) = "T/C"
    xlHoja1.Cells(lnFila + 5, 10) = rs!nValFijo
    TpoCambio = rs!nValFijo 'JUCS
    
    cJudicial = gColocEstRecVigJud & ", " & gColocEstRecVigCast & ", " & gColocEstSolic & ", " & gColocEstSug & ", " & gColocEstRetirado & ", " & gColocEstRech
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue

If psMoneda = "1" Then
    xlHoja1.PageSetup.Zoom = 80
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Reporte 6D", 10)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 6) = "MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 9) = "CALCULO SPREAD" 'JUCS 20170208

    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "MONTO DESEMBOLSADO (en nuevos soles)  " 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "MONTO DESEMBOLSADO (en " & StrConv(gcPEN_PLURAL, vbLowerCase) & ")  "    'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
    xlHoja1.Cells(lnFila + 1, 6) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
    xlHoja1.Cells(lnFila + 1, 7) = "MONTO DESEMBOLSADO (en Dolares americanos)  "
    xlHoja1.Cells(lnFila + 1, 8) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
    xlHoja1.Cells(lnFila + 1, 9) = "MN" 'JUCS 20170208
    xlHoja1.Cells(lnFila + 1, 10) = "ME" 'JUCS 20170208
       
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 10
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 17
    xlHoja1.Range("F1").ColumnWidth = 13
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 12
    xlHoja1.Range("I1").ColumnWidth = 18 'JUCS 20170208
    xlHoja1.Range("J1").ColumnWidth = 18 'JUCS 20170208
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 10)).MergeCells = True 'JUCS 20170208
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
'declare @fecha char(8),@moneda char(1)
'set @fecha='20071031'
'set @moneda='1'
'***BRGO**2010/09/09*********************************************************************
sSql = " SELECT  r.nRangoCod,"
sSql = sSql & "    r.nRangoDes,"
sSql = sSql & "     r.cProdCod Prod,"
sSql = sSql & "     round(SUM(C.nTasaInt)/count(*),2) nTasaInt_Mes,"
sSql = sSql & "     round(SUM(C.ntasCosEfeAnu)/count(*),2) nTCEA,"
sSql = sSql & "     (SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol)/sum(col.nMontoCol)) nTasaInt_Anio, "
sSql = sSql & "     Sum(C.nMontoDesemb) nSaldo,"
sSql = sSql & "     round(SUM(C.nGasto)/count(*),2) PromedioGasto"
sSql = sSql & " FROM " & sservidorconsolidada & "Rangox r"
sSql = sSql & " LEFT JOIN (SELECT   W.cCtaCod, W.cTpoCredCod, "
sSql = sSql & "             W.nTasaInt,"
sSql = sSql & "             isnull(cc.ntasCosEfeAnu,0) ntasCosEfeAnu,"
sSql = sSql & "             W.nMontoDesemb,"
sSql = sSql & "             W.nCuotasApr * CASE WHEN W.nPlazoApr = 0 THEN 30 ELSE W.NPLAZOAPR END nDias,"
sSql = sSql & "             0 as  nGasto "
sSql = sSql & "         FROM " & sservidorconsolidada & "CreditoConsolTotal W"
sSql = sSql & "         left JOIN ColocacCred cc on W.cctacod=cc.cctacod"
sSql = sSql & "         WHERE datediff(m,W.dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 "
sSql = sSql & "         and SubString(W.cCtaCod,9,1) = " & psMoneda & " and W.nPrdEstado"
sSql = sSql & "         IN (2020, 2021, 2022, 2030, 2031, 2032, 2101, 2104, 2106, 2107, 2022, 2092, 2201, 2205)"
'ALPA 20110706*****************************
If psConsol = "N" Then
    sSql = sSql & "      And W.cAgeCodAct in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
'******************************************
sSql = sSql & "     ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin"
sSql = sSql & "     and r.cProdCod = SubString(C.cTpoCredCod,1,1)"
sSql = sSql & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin"
sSql = sSql & "     And ((r.bPrdIn = 1"
sSql = sSql & "         And C.cTpoCredCod In (Select cPrdIn From " & sservidorconsolidada & "RangoDet1 RD"
sSql = sSql & "                         Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                         And RD.nRangoCod = r.nRangoCod))"
sSql = sSql & "         Or (r.bPrdOut = 1 And C.cTpoCredCod "
sSql = sSql & "             Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet1 RD"
sSql = sSql & "                     Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                     And RD.nRangoCod = r.nRangoCod)))"
sSql = sSql & "left join colocaciones col on (C.cCtaCod=col.cCtaCod) "
sSql = sSql & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda
sSql = sSql & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod"
sSql = sSql & " ORDER BY nRangoCod"
'***ALPA**END****************************************************************************
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    nTotalMN = 0 'JUCS 20170209
    nTotSaldoMN = 0
    nTotalME = 0
    nTotSaldoME = 0 'JUCS 20170209
    
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
          
            xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 3)) = rs!nSaldo 'rs!PromedioGasto
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 3)) = rs!nTCEA  'rs!nSaldo
            '******JUCS****2017/02/08**********************************************************************************
            If psMoneda = 1 Then
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalMN = nSaldoMN * nTaeMN
            nTotalMN = nTotalMN + nSubTotalMN
            nTotSaldoMN = nTotSaldoMN + nSaldoMN
            xlHoja1.Cells(lnFila, 9) = nSubTotalMN
            Else
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalME = nSaldoMN * nTaeMN
            nTotalME = nTotalME + nSubTotalME
            nTotSaldoME = nTotSaldoME + nSaldoMN
            xlHoja1.Cells(lnFila, 10) = nSubTotalME
            End If
            '******JUCS****2017/02/08**********************************************************************************
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00"
    End If
    
    '******JUCS****2017/02/08**********************************************************************************
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 9), xlHoja1.Cells(lnFila, 10)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 9), xlHoja1.Cells(lnFila, 10)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 9), xlHoja1.Cells(lnFila, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 9), xlHoja1.Cells(lnFila, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 9), xlHoja1.Cells(lnFila, 10)).NumberFormat = "#,##0.00"
    End If
    If psMoneda = "1" Then 'jucs 20170208
       xlHoja1.Cells(lnFila + 1, 1) = "TOTAL :"
       xlHoja1.Cells(lnFila + 1, 9) = nTotalMN
       xlHoja1.Cells(lnFila + 1, 4) = nTotSaldoMN
    Else
       xlHoja1.Cells(lnFila + 1, 10) = nTotalME
       xlHoja1.Cells(lnFila + 1, 7) = nTotSaldoME
    End If  'jucs 20170202
    
     'Calclulo Final y llenado en celdas de Spread Financiero
    If psMoneda = "1" Then
       nGuardaMN = nTotalMN
       nGuardaTotMN = nTotSaldoMN
       
       If nGuardaTotMN = 0 Then
       TAMN = 0
       Else
       TAMN = nGuardaMN / nGuardaTotMN
       End If
       
       ResultSpread(6) = TAMN
       ReDim Preserve ResultSpread(8)
    Else
        nGuardaMN = nGuardaMN
        nGuardaTotMN = nGuardaTotMN
        SumTotSpread = nGuardaMN + nTotalME * TpoCambio
        SumTotRep6A = nGuardaTotMN + nTotSaldoME * TpoCambio
        
        If nTotSaldoME = 0 Then
        TAME = 0
        Else
        TAME = nTotalME / nTotSaldoME
        End If
        If SumTotRep6A = 0 Then
        TAO = 0
        Else
        TAO = SumTotSpread / SumTotRep6A
        End If
        
        ResultSpread(7) = TAME
        ReDim Preserve ResultSpread(8)
        
        ResultSpread(8) = TAO
        ReDim Preserve ResultSpread(8)
        
        xlHoja1.Cells(lnFila + 1, 11) = SumTotSpread
        xlHoja1.Cells(lnFila + 1, 12) = SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 1) = "SPREAD FINANCIERO"
'        xlHoja1.Cells(lnFila + 2, 9) = nGuardaMN / nGuardaTotMN
'        xlHoja1.Cells(lnFila + 2, 10) = nTotalME / nTotSaldoME
'        xlHoja1.Cells(lnFila + 2, 11) = SumTotSpread / SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 9) = TAMN
        xlHoja1.Cells(lnFila + 2, 10) = TAME
        xlHoja1.Cells(lnFila + 2, 11) = TAO
     End If
     '******JUCS****2017/02/08**********************************************************************************
     
    xlHoja1.Range("C17:H17").Font.Bold = True
    xlHoja1.Range("C20:H20").Font.Bold = True
    xlHoja1.Range("C23:H23").Font.Bold = True
    xlHoja1.Range("C26:H26").Font.Bold = True
    xlHoja1.Range("C29:H29").Font.Bold = True
    xlHoja1.Range("C45:H45").Font.Bold = True
    xlHoja1.Range("C48:H48").Font.Bold = True
    xlHoja1.Range("A67:L68").Font.Bold = True
    
    xlHoja1.Range("A67:L67").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A68:L68").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A69:L69").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A67:L67").NumberFormat = "#,##0.00"
    xlHoja1.Range("A68:L68").NumberFormat = "#,##0.00"
   
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
''peac 20071121

Private Sub Genera6E1(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
    
     sSql = "SELECT nValFijo " 'JUCS 20170203
     sSql = sSql & "FROM TipoCambio "
     sSql = sSql & "WHERE DATEDIFF(d,dFecCamb,dateadd(day,1,('" & Format(pdFecha, gsFormatoFecha) & "'))) =0"
    
    Set rs = oCon.CargaRecordSet(sSql)
    xlHoja1.Cells(lnFila + 5, 7) = "T/C"
    xlHoja1.Cells(lnFila + 5, 8) = rs!nValFijo
    TpoCambio = rs!nValFijo 'JUCS
        
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
   
   lsFechaIni = "01/" & Mid(CStr(pdFecha), 4, 2) & "/" & Mid(CStr(pdFecha), 7, 4)
   lsFechaFin = CStr(pdFecha)

If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Reporte 6E", 8)
    xlHoja1.Range("A1").ColumnWidth = 19
    xlHoja1.Range("B1").ColumnWidth = 18
    xlHoja1.Range("C1").ColumnWidth = 11
    xlHoja1.Range("D1").ColumnWidth = 13
    xlHoja1.Range("E1").ColumnWidth = 11
    xlHoja1.Range("F1").ColumnWidth = 14
    xlHoja1.Range("G1").ColumnWidth = 18 'JUCS 20170208
    xlHoja1.Range("H1").ColumnWidth = 18 'JUCS 20170208
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 7) = "CALCULO SPREAD" 'JUCS 20170208
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "MONTO RECIBIDO 2/(en nuevos soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "MONTO RECIBIDO 2/                                                      (en " & StrConv(gcPEN_PLURAL, vbLowerCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    xlHoja1.Cells(lnFila + 1, 6) = "MONTO RECIBIDO 2/                                                      (en dólares de N.A.)"
    xlHoja1.Cells(lnFila + 1, 7) = "MN" 'JUCS 20170208
    xlHoja1.Cells(lnFila + 1, 8) = "ME" 'JUCS 20170208
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).MergeCells = True 'JUCS 20170208
Else
    lnFila = 10
End If

    oBarra.Progress 1, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2

'DECLARE @INI CHAR(8), @FIN CHAR(8),@MONEDA CHAR(1)
'@INI='20071001'
'@FIN='20071031'
'@MONEDA='1'

Set oConecta = New DConecta
oConecta.AbreConexion
'***ALPA**2008/06/03*********************************************************************
sSql = " SELECT  r.cProdCod Prod,"
sSql = sSql & "     r.nRangoCod,"
sSql = sSql & "     r.nRangoDes,"
sSql = sSql & "     nTasaInt_Anio,"
sSql = sSql & "     SUM(nSaldo) nSaldo"
sSql = sSql & " FROM DBConsolidada..Rango1 R"
sSql = sSql & " LEFT JOIN ( SELECT 31 nRango,-10 nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN sum(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         from Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..CTSConsol ah on ah.cctacod = MVC.cCtaCod"
sSql = sSql & "         where MVC.cOpeCod like '220[1-2]%' and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and  '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag =0 and substring(MVC.cctacod,6,3) = '234' and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
If psConsol = "N" Then
    sSql = sSql & "         and SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "     Union"
sSql = sSql & "     SELECT 2 nRango,-5 nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN sum(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         FROM  Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..ahorrocConsol ah on ah.cctacod = MVC.cctacod"
sSql = sSql & "         where (MVC.cOpeCod like '2001%' or MVC.cOpeCod = '100102')"
sSql = sSql & "         And substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and (nMovflag =0 or nMovflag <> 1 )"
sSql = sSql & "         and substring(MVC.cctacod,6,3) = '232' and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
If psConsol = "N" Then
    sSql = sSql & "         and SubString(isnull(MVC.cCtaCod,'10901'),4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=1"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and r1.nRangoCod between 14 and 18"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=2"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 20 and 24"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria not in (1,2)"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 26 and 30"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni) Dat"
sSql = sSql & "         ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin and dat.nRango=r.nRangoCod"
sSql = sSql & " WHERE cTipoAnx = 'E' and r.nRangoCod <> 32"
sSql = sSql & " GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio"
sSql = sSql & " Union"
sSql = sSql & " SELECT  Prod,"
sSql = sSql & "     nRangoCod,"
sSql = sSql & "     nRangoDes,"
sSql = sSql & "     SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
sSql = sSql & "     SUM(nSaldo) As nSaldo"
sSql = sSql & "     From  (SELECT r.cProdCod Prod,"
sSql = sSql & "             r.nRangoCod,"
sSql = sSql & "             r.nRangoDes,"
sSql = sSql & "             ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
sSql = sSql & "             ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
sSql = sSql & "         FROM DBConsolidada..Rango1 R"
sSql = sSql & "         JOIN (SELECT  -1 nDias,"
sSql = sSql & "                 ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual,"
'sSql = sSql & "                  ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "                  SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/SUM(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "                 CASE WHEN SUM(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/SUM(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "                 SUM(nSaldoContable) nSaldo"
sSql = sSql & "             FROM Mov MV"
sSql = sSql & "             INNER JOIN  MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "             left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod"
sSql = sSql & "             where MVC.cOpeCod LIKE '210[16]%'"
sSql = sSql & "             and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "             and mv.nmovflag=0 and substring(MVC.cctacod,6,3) = '233'"
sSql = sSql & "             and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
sSql = sSql & "             and PF.cCtaCod in (select P.cCtaCod"
sSql = sSql & "                         from DBConsolidada..PlazoFijoConsol P"
sSql = sSql & "                         join DBConsolidada..ProductoBloqueosConsol PB on P.cCtaCod=PB.cCtaCod"
sSql = sSql & "                         where nBlqMotivo=3 and substring(p.cctacod,6,3)='233'"
If psConsol = "N" Then
    sSql = sSql & "             And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "                         and nEstCtaPF in(1100,1200))) Dat"
sSql = sSql & "                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
sSql = sSql & "         WHERE cTipoAnx = 'E' AND nRangoCod = 32"
sSql = sSql & "         GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio)X"
sSql = sSql & " GROUP BY Prod, nRangoCod, nRangoDes"
sSql = sSql & " ORDER BY Prod, nRangoCod, nRangoDes" 'EJVG20131202
'***ALPA**END****************************************************************************
    
    lsProd = "": lnFilaIni = lnFila
    
    If gbBitCentral = True Then
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set rs = oCon.CargaRecordSet(sSql)
    End If
    oBarra.Progress 2, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    
    nTotalMN = 0
    nTotSaldoMN = 0
    nTotalME = 0
    nTotSaldoME = 0
     
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
            '******JUCS****2017/02/08**********************************************************************************
            If psMoneda = 1 Then
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalMN = nSaldoMN * nTaeMN
            nTotalMN = nTotalMN + nSubTotalMN
           
            xlHoja1.Cells(lnFila, 7) = nSubTotalMN
            Else
            nTaeMN = rs!nTasaInt_Anio
            nSaldoMN = rs!nSaldo
            nSubTotalME = nSaldoMN * nTaeMN
            nTotalME = nTotalME + nSubTotalME
            
            xlHoja1.Cells(lnFila, 8) = nSubTotalME
            End If
            '******JUCS****2017/02/08**********************************************************************************
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    '******JUCS****2017/02/08**********************************************************************************
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 7), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00"
    End If
    If psMoneda = "1" Then
       xlHoja1.Cells(lnFila + 1, 1) = "TOTAL :"
       xlHoja1.Cells(lnFila + 1, 7) = nTotalMN
    Else
       xlHoja1.Cells(lnFila + 1, 8) = nTotalME
    End If
    
    If psMoneda = "1" Then
       nGuardaMN = nTotalMN
       'nGuardaTotMN = nTotSaldoMN
    Else
        nGuardaMN = nGuardaMN
        'nGuardaTotMN = nGuardaTotMN
        SumTotSpread = nGuardaMN + nTotalME * TpoCambio
        'SumTotRep6A = nGuardaTotMN + nTotSaldoME * TpoCambio
        xlHoja1.Cells(lnFila + 1, 9) = SumTotSpread
        'xlHoja1.Cells(lnFila + 1, 10) = SumTotRep6A
        xlHoja1.Cells(lnFila + 2, 1) = "SPREAD FINANCIERO"
       ' xlHoja1.Cells(lnFila + 2, 7) = nGuardaMN / nGuardaTotMN
       ' xlHoja1.Cells(lnFila + 2, 8) = nTotalME / nTotSaldoME
       ' xlHoja1.Cells(lnFila + 2, 9) = SumTotSpread / SumTotRep6A
     End If
     '******JUCS****2017/02/08**********************************************************************************
            
    xlHoja1.Range("C24").FormulaR1C1 = "=+((R25C3*R25C4)+(R26C3*R26C4)+(R27C3*R27C4)+(R28C3*R28C4)+(R29C3*R29C4))/IF(SUM(R25C4:R29C4)=0,1,SUM(R25C4:R29C4))"
    xlHoja1.Range("D24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E24").FormulaR1C1 = "=+((R25C5*R25C6)+(R26C5*R26C6)+(R27C5*R27C6)+(R28C5*R28C6)+(R29C5*R29C6))/IF(SUM(R25C6:R29C6)=0,1,SUM(R25C6:R29C6))"
    xlHoja1.Range("F24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C30").FormulaR1C1 = "=+((R31C3*R31C4)+(R32C3*R32C4)+(R33C3*R33C4)+(R34C3*R34C4)+(R35C3*R35C4))/IF(SUM(R31C4:R35C4)=0,1,SUM(R31C4:R35C4))"
    xlHoja1.Range("D30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E30").FormulaR1C1 = "=+((R31C5*R31C6)+(R32C5*R32C6)+(R33C5*R33C6)+(R34C5*R34C6)+(R35C5*R35C6))/IF(SUM(R31C6:R35C6)=0,1,SUM(R31C6:R35C6))"
    xlHoja1.Range("F30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C36").FormulaR1C1 = "=+((R37C3*R37C4)+(R38C3*R38C4)+(R39C3*R39C4)+(R40C3*R40C4)+(R41C3*R41C4))/IF(SUM(R37C4:R41C4)=0,1,SUM(R37C4:R41C4))"
    xlHoja1.Range("D36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E36").FormulaR1C1 = "=+((R37C5*R37C6)+(R38C5*R38C6)+(R39C5*R39C6)+(R40C5*R40C6)+(R41C5*R41C6))/IF(SUM(R37C6:R41C6)=0,1,SUM(R37C6:R41C6))"
    xlHoja1.Range("F36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    
    'Calclulo Final y llenado en celdas de Spread Financiero
    xlHoja1.Range("D45").FormulaR1C1 = "=+(R13C4+R24C4+R30C4+R36C4+R42C4+R43C4)"
    xlHoja1.Range("F45").FormulaR1C1 = "=+(R13C6+R24C6+R30C6+R36C6+R42C6+R43C6)"
    xlHoja1.Range("F45").FormulaR1C1 = "=+(R13C6+R24C6+R30C6+R36C6+R42C6+R43C6)"
    xlHoja1.Range("J45").FormulaR1C1 = "=+((R45C4)+(R45C6)*(R5C8))"
    xlHoja1.Range("G46").FormulaR1C1 = "=+(R45C7/R45C4)"
    xlHoja1.Range("H46").FormulaR1C1 = "=+(R45C8/R45C6)"
    xlHoja1.Range("I46").FormulaR1C1 = "=+(R45C9/R45C10)"
    
    xlHoja1.Range("C24:H24").Font.Bold = True
    xlHoja1.Range("C30:H30").Font.Bold = True
    xlHoja1.Range("C36:H36").Font.Bold = True
    xlHoja1.Range("A45:J46").Font.Bold = True
    xlHoja1.Range("A45:J45").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A46:J46").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A47:J47").Borders(xlEdgeTop).LineStyle = True
    xlHoja1.Range("A45:J45").NumberFormat = "#,##0.00"
    xlHoja1.Range("A46:J46").NumberFormat = "#,##0.00"
   
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
'**************inicio JUCS**********2017/03/13*********************************************************************
Private Sub GeneraSpread(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim I  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim sSqlTemp As String
    Dim rs As New ADODB.Recordset
    Dim cPigno As String
    Dim cVigente As String

   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "REPORTE SPREAD FINANCIERO", "Cargando Datos", "", vbBlue
    
If psMoneda = "1" Then
    oBarra.Progress 1, "REPORTE SPREAD FINANCIERO", "Tasas Act y Pas sobre saldos", "", vbBlue
    'lnFila = lnFila + 2
    
    lnFila = CabeceraReporteSpread(pdFecha, "CALCULO SPREAD FINANCIERO", "Spread", 9)
    xlHoja1.Cells(lnFila, 1) = "TASAS DE INTERSES ACTIVAS Y PASIVAS SOBRE SALDOS"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).MergeCells = True
    lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 1) = "FECHA"
    xlHoja1.Cells(lnFila, 1) = "ACTIVAS"
    xlHoja1.Cells(lnFila, 4) = "PASIVAS"
    xlHoja1.Cells(lnFila, 7) = "SPREAD"

    xlHoja1.Cells(lnFila + 1, 1) = "TAMN"
    xlHoja1.Cells(lnFila + 1, 2) = "TAME"
    xlHoja1.Cells(lnFila + 1, 3) = "TAO"
    xlHoja1.Cells(lnFila + 1, 4) = "TPMN"
    xlHoja1.Cells(lnFila + 1, 5) = "TPME"
    xlHoja1.Cells(lnFila + 1, 6) = "TAO"
    
    xlHoja1.Cells(lnFila + 1, 7) = "SPREAD - MN"
    xlHoja1.Cells(lnFila + 1, 8) = "SPREAD - ME"
    xlHoja1.Cells(lnFila + 1, 9) = "SPREAD - O"
    
    xlHoja1.Range("A1").ColumnWidth = 18
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 14
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 14
    xlHoja1.Range("G1").ColumnWidth = 14
    xlHoja1.Range("H1").ColumnWidth = 14
    xlHoja1.Range("I1").ColumnWidth = 14
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 3)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 9)).MergeCells = True
    
    lnFila = lnFila + 1
    'para combinar resultado spread
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila + 1, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila + 1, 9)).MergeCells = True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "de <REPORTE 6A>"
    xlHoja1.Cells(lnFila, 4) = "de <REPORTE 6B>"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 3)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    
    'Calculo Final Spread
    xlHoja1.Range("G14").FormulaR1C1 = "=(R14C1)-(R14C4)"
    xlHoja1.Range("H14").FormulaR1C1 = "=(R14C2)-(R14C5)"
    xlHoja1.Range("I14").FormulaR1C1 = "=(R14C3)-(R14C6)"
    lnFila = lnFila + 1
    
'    xlHoja1.Cells(lnFila, 1) = "=Rep_6A!G45"
'    xlHoja1.Cells(lnFila, 2) = "=Rep_6A!H45"
'    xlHoja1.Cells(lnFila, 3) = "=Rep_6A!I45"
'
'    xlHoja1.Cells(lnFila, 4) = "=Rep_6B!G39"
'    xlHoja1.Cells(lnFila, 5) = "=Rep_6B!H39"
'    xlHoja1.Cells(lnFila, 6) = "=Rep_6B!I39"

    For I = 0 To 5
         Select Case I
                Case 0
                xlHoja1.Cells(lnFila, 1) = ResultSpread(I)
                Case 1
                xlHoja1.Cells(lnFila, 2) = ResultSpread(I)
                Case 2
                xlHoja1.Cells(lnFila, 3) = ResultSpread(I)
                Case 3
                xlHoja1.Cells(lnFila, 4) = ResultSpread(I)
                Case 4
                xlHoja1.Cells(lnFila, 5) = ResultSpread(I)
                Case 5
                xlHoja1.Cells(lnFila, 6) = ResultSpread(I)
         End Select
    Next I
    xlHoja1.Range("A14:i14").NumberFormat = "#,##0.00"
    
    
    oBarra.Progress 2, "REPORTE SPREAD FINANCIERO", "Tasas Act y Pas sobre saldos", "", vbBlue
    
    'cuadro N° 2****** Tasas de Interes activas y pasivas de Operaciones Diarias*************
    oBarra.Progress 1, "REPORTE SPREAD FINANCIERO", "Tasas Act y Pas Ope Diarias", "", vbBlue
    
    lnFila = lnFila + 5
    lnFila = SegundaCabeceraSpread(pdFecha, "", "", 9)
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).MergeCells = True
    xlHoja1.Cells(lnFila, 1) = "TASAS DE INTERSES ACTIVAS Y PASIVAS DE OPERACIONES DIARIAS"
    lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 1) = "FECHA"
    xlHoja1.Cells(lnFila, 1) = "ACTIVAS"
    xlHoja1.Cells(lnFila, 4) = "PASIVAS"
    xlHoja1.Cells(lnFila, 7) = "SPREAD"

    xlHoja1.Cells(lnFila + 1, 1) = "TAMN"
    xlHoja1.Cells(lnFila + 1, 2) = "TAME"
    xlHoja1.Cells(lnFila + 1, 3) = "TAO"
    xlHoja1.Cells(lnFila + 1, 4) = "TPMN"
    xlHoja1.Cells(lnFila + 1, 5) = "TPME"
    xlHoja1.Cells(lnFila + 1, 6) = "TAO"
    
    xlHoja1.Cells(lnFila + 1, 7) = "SPREAD - MN"
    xlHoja1.Cells(lnFila + 1, 8) = "SPREAD - ME"
    xlHoja1.Cells(lnFila + 1, 9) = "SPREAD - O"
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 3)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 9)).MergeCells = True
    
    lnFila = lnFila + 1
    'para conbinar resultado spread
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila + 1, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila + 1, 9)).MergeCells = True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "de <REPORTE 6D>"
    xlHoja1.Cells(lnFila, 4) = "de <REPORTE 6E>"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 3)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    'Calculo Final Spread
    xlHoja1.Range("G22").FormulaR1C1 = "=(R22C1)-(R22C4)"
    xlHoja1.Range("H22").FormulaR1C1 = "=(R22C2)-(R22C5)"
    xlHoja1.Range("I22").FormulaR1C1 = "=(R22C3)-(R22C6)"
    
    lnFila = lnFila + 1

   For I = 6 To 8
        Select Case I
               Case 6
               xlHoja1.Cells(lnFila, 1) = ResultSpread(I)
               Case 7
               xlHoja1.Cells(lnFila, 2) = ResultSpread(I)
               Case 8
               xlHoja1.Cells(lnFila, 3) = ResultSpread(I)
        End Select
   Next I
   
'  xlHoja1.Cells(lnFila, 1) = "='Rep_6D (nuevo)'!I68"
'  xlHoja1.Cells(lnFila, 2) = "='Rep_6D (nuevo)'!J68"
'  xlHoja1.Cells(lnFila, 3) = "='Rep_6D (nuevo)'!K68"
  
  xlHoja1.Cells(lnFila, 4) = "='Rep_6E (nuevo)'!G46"
  xlHoja1.Cells(lnFila, 5) = "='Rep_6E (nuevo)'!H46"
  xlHoja1.Cells(lnFila, 6) = "='Rep_6E (nuevo)'!I46"
  
  xlHoja1.Range("A22:I22").NumberFormat = "#,##0.00"
Else
    lnFila = 23
End If

    oBarra.Progress 2, "REPORTE SPREAD FINANCIERO", "Datos insertados", "", vbRed
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
'***************fin JUCS**********2017/03/13*********************************************************************
Private Function CabeceraReporteSpread(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer) As Integer
Dim lnFila As Integer
    xlHoja1.Range("A1:R100").Font.Size = 8

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
    
    lnFila = 1
    xlHoja1.Cells(lnFila, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 5) = psTitulo
    'lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 9) = psReporte
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 1) = "EMPRESA : " & gsNomCmac:
    xlHoja1.Cells(lnFila + 1, 1) = "Fecha : AL " & Format(pdFecha, "dd mmmm yyyy")
    
    'lnFila = lnFila + 1
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).MergeCells = True
    
    lnFila = lnFila + 3
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Font.Bold = True
    CabeceraReporteSpread = lnFila
    
End Function
Private Function SegundaCabeceraSpread(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer) As Integer
 Dim lnFila As Integer
 'xlHoja1.Range("A1:R100").Font.Size = 18
 'xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
 lnFila = 18
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 4, pnCols)).Font.Bold = True
    SegundaCabeceraSpread = lnFila
End Function

Private Sub Form_Load()
Set oCon = New DConecta
oCon.AbreConexion
CentraForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub
