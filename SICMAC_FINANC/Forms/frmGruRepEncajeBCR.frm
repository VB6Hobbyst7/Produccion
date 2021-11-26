VERSION 5.00
Begin VB.Form frmGruRepEncajeBCR 
   Caption         =   "Encaje BCR"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2475
   LinkTopic       =   "Form2"
   ScaleHeight     =   705
   ScaleWidth      =   2475
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmGruRepEncajeBCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsMoneda As String
Dim lsCodOpe As String
Dim oRep As DRepFormula
Dim lsSaldos() As String
Dim Ctamas, Ctamenos As String
Dim colmas, colmenos As String
Dim ldFecIni, ldFecFin As Date
Dim ldFechaAnt As String
'Nombre del Archivo
Dim lsArchivo As String
Dim lbExcelOpen As Boolean
'Variables de excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHojaP As Excel.Worksheet
'Anexo 2
Dim lnTCF  As Currency
Dim lnTCFF As Currency
Dim lNumColCAJAMERep1 As Integer 'PASI20150310

Public Sub ImprimeEncajeBCR(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, ByVal pnlnTCF As Double, ByVal plnTCFF As Double)
   lsMoneda = Mid(psOpeCod, 3, 1)
   ldFecIni = pdFecIni
   ldFecFin = pdFecFin
   'lsArchivo = App.path & "\Spooler\RepEncajeBCR_" & Left(Format(pdFecIni, gsFormatoMovFecha), 6) & IIf(lsMoneda = "1", "MN", "ME") & ".XLS"
   
    Select Case psOpeCod
        Case "761206", "762206" 'Reporte 1: Obligaciones Sujetas a Encaje
            lsArchivo = App.path & "\Spooler\C1" & Format(DatePart("M", pdFecIni), "00") & "01" & IIf(lsMoneda = "1", "MN", "ME") & ".xlsx"
        Case "761207", "762207"
            lsArchivo = App.path & "\Spooler\C1" & Format(DatePart("M", pdFecIni), "00") & "02" & IIf(lsMoneda = "1", "MN", "ME") & ".xlsx"
        Case "761208", "762208"
            lsArchivo = App.path & "\Spooler\C1" & Format(DatePart("M", pdFecIni), "00") & "03" & IIf(lsMoneda = "1", "MN", "ME") & ".xlsx"
        Case "761209", "762209"
            lsArchivo = App.path & "\Spooler\C1" & Format(DatePart("M", pdFecIni), "00") & "04" & IIf(lsMoneda = "1", "MN", "ME") & ".xlsx"
   End Select
   
   lNumColCAJAMERep1 = 0
   lsCodOpe = psOpeCod
    If lsMoneda = "2" Then
        Dim oTC As nTipoCambio
        'Modificado PASIERS1712014
        'lnTCFF = pnlnTCF
        lnTCF = pnlnTCF
        lnTCFF = plnTCFF
        'END PASI
    End If
    If Not ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False) Then
        Exit Sub
    End If
    lbExcelOpen = True
        Select Case psOpeCod
            Case "761206", "762206" 'Reporte 1: Obligaciones Sujetas a Encaje
                ImprimeObligacionSujetaEnc
            Case "761207", "762207"
                ImprimeObligNoSujEncIfisPais
            Case "761208", "762208"
                ImprimeObligIFISPais
            Case "761209", "762209"
                ImprimeOtrasOblig
        End Select
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    CargaArchivo lsArchivo, App.path & "\SPOOLER"
    lbExcelOpen = False
End Sub
Public Sub ImprimeOtrasOblig()
    CabeceraReporte
    ContenidoOtrasOblig
End Sub
Private Sub ContenidoOtrasOblig()
    CabeceraReporte4
    ContenidoSaldoRep4
End Sub
Private Sub ContenidoSaldoRep4()
    Set oRep = New DRepFormula
    Dim ldFecha As Date
    Dim rsrep4 As ADODB.Recordset
    Set rsrep4 = New ADODB.Recordset
    Dim lncol, lnfil, i, j As Integer
    Dim lnTotalCol As Integer
    Dim lnSaldoTot As Currency 'para obtener los totales por columna
    Dim lnSaldo As Currency
    
    ldFecha = ldFecIni
    Set rsrep4 = oRep.ObtenerSubColRep4FormulaBCR(CInt(lsMoneda))
    lnTotalCol = rsrep4.RecordCount
    ReDim lsSaldos(1 To lnTotalCol, 0 To 0)
    lncol = 0
    lnfil = 0
    lnSaldoTot = 0
    Do While ldFecha <= ldFecFin
        lnfil = lnfil + 1
        ReDim Preserve lsSaldos(1 To lnTotalCol, lnfil)
        Do While Not rsrep4.EOF
            lncol = lncol + 1
            If rsrep4!nValor = 1 Then
                lnSaldo = DevolverSaldoFormulaRep1(rsrep4!cvalor, ldFecha)
                lnSaldoTot = lnSaldoTot + lnSaldo
                lsSaldos(lncol, lnfil) = lnSaldo
            Else
                lsSaldos(lncol, lnfil) = lnSaldoTot
                lnSaldo = 0
                lnSaldoTot = 0
            End If
            rsrep4.MoveNext
        Loop
        rsrep4.MoveFirst
        lnSaldo = 0
        lncol = 0
        lnSaldoTot = 0
        ldFecha = DateAdd("d", 1, ldFecha)
    Loop
    Set oRep = Nothing
    
    'Ingresa los datos al excel
    lnfil = 21
    lncol = 0
    
    Dim lnSaldoCelda() As Currency
    ReDim lnSaldoCelda(1 To UBound(lsSaldos, 1), 1)
    
    For i = 1 To UBound(lnSaldoCelda, 1)
        lnSaldoCelda(i, 1) = 0
    Next i
    
    For i = 1 To UBound(lsSaldos, 2)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, lncol) = i
        For j = 1 To UBound(lsSaldos, 1)
            lncol = lncol + 1
            lnSaldoCelda(j, 1) = lnSaldoCelda(j, 1) + lsSaldos(j, i)
            xlHoja1.Cells(lnfil, lncol) = Format(lsSaldos(j, i), "#,#0.00")
        Next j
        lncol = 0
        lnfil = lnfil + 1
    Next i
    lncol = 1
    For i = 1 To UBound(lnSaldoCelda, 1)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, 1) = "TOTAL"
        xlHoja1.Cells(lnfil, lncol) = Format(lnSaldoCelda(i, 1), "#,#0.00")
    Next i
End Sub
Private Sub CabeceraReporte4()
    Set oRep = New DRepFormula
    Dim lx, lreg, lfil, lcol, lz, ly, i As Integer
    Dim ltemp As Boolean
    Dim lcolumna As String
    Dim oValoresCab As ADODB.Recordset
    Dim sCadena As String
    Set oValoresCab = New ADODB.Recordset
    Set oValoresCab = oRep.ObtenerCabeceraRep4FormulaBCR(CInt(lsMoneda))
    Dim ldFecha As String
    Dim lnPlaz As Double
    lx = 2
    ltemp = False
    lfil = 13
    lcol = 2
    lz = 0
    ly = 1
    xlHoja1.Cells(lfil + 1, 1) = "COD. OPERACIÓN"
    xlHoja1.Cells(lfil + 4, 1) = "DIA"
    xlHoja1.Cells(lfil + 5, 1) = "Fecha de Inicio"
    xlHoja1.Cells(lfil + 6, 1) = "Fecha de Vcto."
    xlHoja1.Cells(lfil + 7, 1) = "PLAZO PROMEDIO"
    
        
    Do While Not oValoresCab.EOF
        If Not ltemp Then
            lreg = oValoresCab!cCodigo
            xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
            xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodigo
            ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
            xlHoja1.Cells(lfil + 5, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
            ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
            xlHoja1.Cells(lfil + 6, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
            lnPlaz = oValoresCab!nPlazoProm
            xlHoja1.Cells(lfil + 7, lx) = IIf(lnPlaz = 0, "", lnPlaz)
            xlHoja1.Cells(lfil + 100, lx) = oValoresCab!cCodSwif
            ltemp = True
            lx = lx + 1
            ly = ly + 1
            lz = lz + 1
            lcolumna = oValoresCab!columna
        Else
            If (oValoresCab!cCodigo = lreg) Or (Mid(oValoresCab!cCodigo, 1, 2) = Mid(lreg, 1, 2)) Then
                xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodigo
                ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                xlHoja1.Cells(lfil + 5, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                xlHoja1.Cells(lfil + 6, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                lnPlaz = oValoresCab!nPlazoProm
                xlHoja1.Cells(lfil + 7, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                xlHoja1.Cells(lfil + 100, lx) = oValoresCab!cCodSwif
                lx = lx + 1
                lz = lz + 1
                lcolumna = oValoresCab!columna
            Else
                xlHoja1.Cells(lfil - 1, ly) = lcolumna
                xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodigo
                ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                xlHoja1.Cells(lfil + 5, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                xlHoja1.Cells(lfil + 6, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                lnPlaz = oValoresCab!nPlazoProm
                xlHoja1.Cells(lfil + 7, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                xlHoja1.Cells(lfil + 100, lx) = oValoresCab!cCodSwif
                lcolumna = oValoresCab!columna
                ly = lx
                lz = 1
                lx = lx + 1
                lreg = oValoresCab!cCodigo
            End If
        End If
        oValoresCab.MoveNext
    Loop
        RSClose oValoresCab
        xlHoja1.Cells(lfil - 1, ly) = lcolumna
        xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
        Set oRep = Nothing
End Sub
Public Sub ImprimeObligIFISPais()
    CabeceraReporte
    ContenidoObligIFISPais
End Sub
Private Sub ContenidoObligIFISPais()
    CabeceraRep3
    ContenidoSaldoRep3
End Sub
Private Sub ContenidoSaldoRep3()
    Set oRep = New DRepFormula
    Dim ldFecha As Date
    Dim rsrep3 As ADODB.Recordset
    Set rsrep3 = New ADODB.Recordset
    Dim rsrep3b As ADODB.Recordset
    Set rsrep3b = New ADODB.Recordset
    Dim lncol, lnfil, i, j As Integer
    Dim lnTotalCol As Integer
    Dim lnSaldoTot As Currency 'para obtener los totales por columna
    Dim lnSaldo As Currency
    Dim lnSaldoTotObl As Currency
    
    ldFecha = ldFecIni
    
    Set rsrep3 = oRep.ObtenerSubColRep3FormulaBCR(2, CInt(lsMoneda))
    Set rsrep3b = oRep.ObtenerSubColRep3FormulaBCR(1, CInt(lsMoneda))
    If Not rsrep3.EOF And Not rsrep3.BOF Then
        lnTotalCol = CInt(rsrep3.RecordCount)
    End If
    If Not rsrep3b.EOF And Not rsrep3b.BOF Then
        lnTotalCol = lnTotalCol + CInt(rsrep3b.RecordCount)
    End If
    
    ReDim lsSaldos(1 To lnTotalCol, 0 To 0)
    lncol = 0
    lnfil = 0
    lnSaldo = 0
    lnSaldoTot = 0
    lnSaldoTotObl = 0
    'Para las Obligaciones no Sujetas a Encaje
    If Not rsrep3.EOF And Not rsrep3.BOF Then
        Do While ldFecha <= ldFecFin
            lnfil = lnfil + 1
            ReDim Preserve lsSaldos(1 To lnTotalCol, lnfil)
            Do While Not rsrep3.EOF
                lncol = lncol + 1
                If rsrep3!nValor = 1 Then
                    lnSaldo = DevolverSaldoFormulaRep2(rsrep3!cvalor, ldFecha, lsMoneda)
                    lnSaldoTot = lnSaldoTot + lnSaldo
                    lsSaldos(lncol, lnfil) = lnSaldo
                ElseIf rsrep3!nValor = 2 Then
                    lsSaldos(lncol, lnfil) = lnSaldoTot
                    lnSaldoTotObl = lnSaldoTotObl + lnSaldoTot
                    lnSaldo = 0
                    lnSaldoTot = 0
                Else
                    lsSaldos(lncol, lnfil) = lnSaldoTotObl
                    lnSaldo = 0
                    lnSaldoTot = 0
                    lnSaldoTotObl = 0
                End If
                rsrep3.MoveNext
            Loop
            rsrep3.MoveFirst
            lnSaldo = 0
            lncol = 0
            lnSaldoTot = 0
            lnSaldoTotObl = 0
            ldFecha = DateAdd("d", 1, ldFecha)
        Loop
    End If
    ldFecha = ldFecIni
    'Para las Obligaciones Sujetas a Encaje
    If Not rsrep3b.EOF And Not rsrep3b.BOF Then
        Do While ldFecha <= ldFecFin
            lnfil = lnfil + 1
            ReDim Preserve lsSaldos(1 To lnTotalCol, lnfil)
            Do While Not rsrep3b.EOF
                lncol = lncol + 1
                If rsrep3b!nValor = 1 Then
                    lnSaldo = DevolverSaldoFormulaRep2(rsrep3b!cvalor, ldFecha, lsMoneda)
                    lnSaldoTot = lnSaldoTot + lnSaldo
                    lsSaldos(lncol, lnfil) = lnSaldo
                ElseIf rsrep3b!nValor = 2 Then
                    lsSaldos(lncol, lnfil) = lnSaldoTot
                    lnSaldoTotObl = lnSaldoTotObl + lnSaldoTot
                    lnSaldo = 0
                    lnSaldoTot = 0
                Else
                    lsSaldos(lncol, lnfil) = lnSaldoTot
                    lnSaldo = 0
                    lnSaldoTot = 0
                    lnSaldoTotObl = 0
                End If
            Loop
            rsrep3.MoveFirst
            lnSaldo = 0
            lncol = 0
            lnSaldoTot = 0
            lnSaldoTotObl = 0
            ldFecha = DateAdd("d", 1, ldFecha)
        Loop
    End If
    
    Set oRep = Nothing
    'Ingresa los Datos al excel
    lnfil = 21
    lncol = 0
    
    Dim lnSaldoCelda() As Currency
    ReDim lnSaldoCelda(1 To UBound(lsSaldos, 1), 1)
    
    For i = 1 To UBound(lnSaldoCelda, 1)
        lnSaldoCelda(i, 1) = 0
    Next i
    
    For i = 1 To UBound(lsSaldos, 2)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, lncol) = i
        For j = 1 To UBound(lsSaldos, 1)
            lncol = lncol + 1
            lnSaldoCelda(j, 1) = lnSaldoCelda(j, 1) + lsSaldos(j, i)
            xlHoja1.Cells(lnfil, lncol) = Format(lsSaldos(j, i), "#,#0.00")
        Next j
        lncol = 0
        lnfil = lnfil + 1
    Next i
    lncol = 1
    For i = 1 To UBound(lnSaldoCelda, 1)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, 1) = "TOTAL"
        xlHoja1.Cells(lnfil, lncol) = Format(lnSaldoCelda(i, 1), "#,#0.00")
    Next i
End Sub
Private Sub CabeceraRep3()
    Set oRep = New DRepFormula
    Dim lx, lreg, lfil, lcol, lz, ly, i As Integer
    Dim ltemp, bvalor As Boolean
    Dim lcolumna As String
    Dim oValoresCab As ADODB.Recordset
    Dim sCadena As String
    Set oValoresCab = New ADODB.Recordset
    Set oValoresCab = oRep.ObtenerCabeceraRep3FormulaBCR(2, CInt(lsMoneda))
    Dim ldFecha As String
    Dim lnPlaz As Double
    Dim m, n As Integer
    lx = 2
    ltemp = False
    lfil = 11
    lcol = 2
    lz = 0
    ly = 1
    n = 1
    m = 1
    'Para obligaciones no sujeta a encaje
    
        xlHoja1.Cells(14, 1) = "Cód. Operación"
        xlHoja1.Cells(15, 1) = "Cód. Swift"
        xlHoja1.Cells(16, 1) = "Destino Financiamiento"
        xlHoja1.Cells(17, 1) = "DIA"
        xlHoja1.Cells(18, 1) = "Fecha Inicio"
        xlHoja1.Cells(19, 1) = "Fecha Vcto."
        xlHoja1.Cells(20, 1) = "PLAZO PROMEDIO"
        
        If Not oValoresCab.EOF And Not oValoresCab.BOF Then
            Do While Not oValoresCab.EOF
                If oValoresCab!nValor = 3 Then
                    xlHoja1.Cells(lfil - 1, lx) = oValoresCab!cDescripcion
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                    xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                    xlHoja1.Cells(lfil - 1, ly) = lcolumna
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                    
                    ly = lx
                    lz = 1
                    lx = lx + 1
                    ltemp = False
                    bvalor = False
                    m = m + 1
                Else
                    If Not ltemp Then
                        lreg = oValoresCab!cCodigo
                        xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                        xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                        xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                        xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                        xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                        ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                        xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                        ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                        xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                        lnPlaz = oValoresCab!nPlazoProm
                        xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz) 'oValoresCab!nPlazoProm
                        ltemp = True
                        lx = lx + 1
                        ly = ly + 1
                        lz = lz + 1
                        lcolumna = oValoresCab!columna
                        m = m + 1
                    Else
                        If (oValoresCab!cCodigo = lreg) Or (Mid(oValoresCab!cCodigo, 1, 3) = Mid(lreg, 1, 3)) Then
                            xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                            xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                            xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                            xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                            xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                            ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            lnPlaz = oValoresCab!nPlazoProm
                            xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                            
                            lx = lx + 1
                            lz = lz + 1
                            lcolumna = oValoresCab!columna
                            m = m + 1
                        Else
                            xlHoja1.Cells(lfil - 1, ly) = lcolumna
                            xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                            xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                            xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                            xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                            xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                            xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                            ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            lnPlaz = oValoresCab!nPlazoProm
                            xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                            lcolumna = oValoresCab!columna
                            ly = lx
                            lz = 1
                            lx = lx + 1
                            lreg = oValoresCab!cCodigo
                            m = m + 1
                        End If
                    End If
                    bvalor = True
                End If
                oValoresCab.MoveNext
            Loop
                RSClose oValoresCab
                If bvalor = True Then
                    xlHoja1.Cells(lfil - 1, ly) = lcolumna
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                End If
            
            xlHoja1.Cells(9, n) = "OBLIGACIONES NO SUJETAS A ENCAJE"
            xlHoja1.Range(xlHoja1.Cells(9, n), xlHoja1.Cells(9, m)).Merge True
            xlHoja1.Cells(8, n) = "ENTIDADES"
            xlHoja1.Range(xlHoja1.Cells(8, n), xlHoja1.Cells(8, m)).Merge True
     End If
     Set oRep = New DRepFormula
     Set oValoresCab = oRep.ObtenerCabeceraRep3FormulaBCR(1, CInt(lsMoneda))
     n = m + 1
     If Not oValoresCab.EOF And oValoresCab.BOF Then
        Do While Not oValoresCab.EOF
           If oValoresCab!nValor = 3 Then
                    xlHoja1.Cells(lfil - 1, lx) = oValoresCab!cDescripcion
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                    xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                    xlHoja1.Cells(lfil - 1, ly) = lcolumna
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                    
                    ly = lx
                    lz = 1
                    lx = lx + 1
                    ltemp = False
                    bvalor = False
                    m = m + 1
            Else
                If Not ltemp Then
                    lreg = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                    xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                    xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                    xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                    ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                    xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                    ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                    xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                    lnPlaz = oValoresCab!nPlazoProm
                    xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                    ltemp = True
                    lx = lx + 1
                    ly = ly + 1
                    lz = lz + 1
                    lcolumna = oValoresCab!columna
                    m = m + 1
                Else
                    If oValoresCab!cCodigo = lreg Then
                            xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                            xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                            xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                            xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                            xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                            ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                            xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                            lnPlaz = oValoresCab!nPlazoProm
                            xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                            
                            lx = lx + 1
                            lz = lz + 1
                            lcolumna = oValoresCab!columna
                            m = m + 1
                    Else
                        xlHoja1.Cells(lfil - 1, ly) = lcolumna
                        xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                        xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                        xlHoja1.Range(xlHoja1.Cells(lfil, lx), xlHoja1.Cells(lfil + 2, lx)).MergeCells = True
                        xlHoja1.Cells(lfil + 3, lx) = oValoresCab!cCodigo
                        xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodSwif
                        xlHoja1.Cells(lfil + 5, lx) = oValoresCab!Destino
                        ldFecha = Format(oValoresCab!dPeriodoDes, "dd/MM/yyyy")
                        xlHoja1.Cells(lfil + 7, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                        ldFecha = Format(oValoresCab!dPeriodoHas, "dd/MM/yyyy")
                        xlHoja1.Cells(lfil + 8, lx) = IIf(ldFecha = "01/01/1900", "", ldFecha)
                        lnPlaz = oValoresCab!nPlazoProm
                        xlHoja1.Cells(lfil + 9, lx) = IIf(lnPlaz = 0, "", lnPlaz)
                        lcolumna = oValoresCab!columna
                        ly = lx
                        lz = 1
                        lx = lx + 1
                        lreg = oValoresCab!cCodigo
                        m = m + 1
                    End If
                End If
                bvalor = True
            End If
            oValoresCab.MoveNext
        Loop
                RSClose oValoresCab
                If bvalor = True Then
                    xlHoja1.Cells(lfil - 1, ly) = lcolumna
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                End If
            
            xlHoja1.Cells(9, n) = "OBLIGACIONES SUJETAS A ENCAJE"
            xlHoja1.Range(xlHoja1.Cells(9, n), xlHoja1.Cells(9, m)).Merge True
            xlHoja1.Cells(8, n) = "ENTIDADES"
            xlHoja1.Range(xlHoja1.Cells(8, n), xlHoja1.Cells(8, m)).Merge True
     End If
End Sub
Public Sub ImprimeObligNoSujEncIfisPais()
    CabeceraReporte
    ContenidoObligNoSujEncIfisPais
End Sub
Private Sub ContenidoObligNoSujEncIfisPais()
    CabeceraRep2
    ContenidoSaldoRep2
End Sub
Private Sub ContenidoSaldoRep2()
    Set oRep = New DRepFormula
    Dim ldFecha As Date
    Dim rsrep2 As ADODB.Recordset
    Set rsrep2 = New ADODB.Recordset
    Dim lncol, lnfil, i, j As Integer
    Dim lnTotalCol As Integer
    Dim lnSaldoTot As Currency 'para obtener los totales por columna
    Dim lnSaldo As Currency
     
    ldFecha = ldFecIni
    Set rsrep2 = oRep.ObtenerSubColRep2FormulaBCR(CInt(lsMoneda))
    lnTotalCol = rsrep2.RecordCount
    ReDim lsSaldos(1 To lnTotalCol, 0 To 0)
    lncol = 0
    lnfil = 0
    lnSaldo = 0
    lnSaldoTot = 0
    Do While ldFecha <= ldFecFin
        lnfil = lnfil + 1
        ReDim Preserve lsSaldos(1 To lnTotalCol, lnfil)
        Do While Not rsrep2.EOF
            lncol = lncol + 1
            If rsrep2!nValor = 1 Then
                    lnSaldo = DevolverSaldoFormulaRep2(rsrep2!cvalor, ldFecha, lsMoneda)
                    lnSaldoTot = lnSaldoTot + lnSaldo
                    lsSaldos(lncol, lnfil) = lnSaldo
            Else
                lsSaldos(lncol, lnfil) = lnSaldoTot
                lnSaldo = 0
                lnSaldoTot = 0
            End If
            rsrep2.MoveNext
        Loop
        rsrep2.MoveFirst
        lnSaldo = 0
        lncol = 0
        lnSaldoTot = 0
        ldFecha = DateAdd("d", 1, ldFecha)
    Loop
    Set oRep = Nothing
    'Ingresar los datos al excel
    lnfil = 14
    lncol = 0
    
    Dim lnSaldoCelda() As Currency
    ReDim lnSaldoCelda(1 To UBound(lsSaldos, 1), 1)
    
    For i = 1 To UBound(lnSaldoCelda, 1)
        lnSaldoCelda(i, 1) = 0
    Next i
    
    For i = 1 To UBound(lsSaldos, 2)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, lncol) = i
        For j = 1 To UBound(lsSaldos, 1)
            lncol = lncol + 1
            lnSaldoCelda(j, 1) = lnSaldoCelda(j, 1) + lsSaldos(j, i)
            xlHoja1.Cells(lnfil, lncol) = Format(lsSaldos(j, i), "#,#0.00")
        Next j
        lncol = 0
        lnfil = lnfil + 1
    Next i
    lncol = 1
    For i = 1 To UBound(lnSaldoCelda, 1)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, 1) = "TOTAL"
        xlHoja1.Cells(lnfil, lncol) = Format(lnSaldoCelda(i, 1), "#,#0.00")
    Next i
End Sub
Private Function DevolverSaldoFormulaRep2(ByVal psValor As String, ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Set oRep = New DRepFormula
    Dim i As Integer
    Dim SaldoMas, SaldoMenos As Currency
    Dim rsSaldo As ADODB.Recordset
    Set rsSaldo = New ADODB.Recordset
    Dim numero As String
    numero = ""
    SaldoMas = 0
    SaldoMenos = 0
    GeneraCuentas (psValor)
    For i = 1 To Len(Ctamas)
        If (Mid(Ctamas, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rsSaldo = oRep.ObtenerRep2SubColCtaSaldo(pdFecha, numero, pnMoneda)
            If pnMoneda = 1 Then
                If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                    SaldoMas = SaldoMas + rsSaldo!nSaldo
                End If
            Else
                If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                    'Modificado PASIERS1382014
                    'SaldoMas = SaldoMas + Round(rsSaldo!nSaldo / lnTCFF, 2)
                    If (pdFecha = ldFecFin) And DateAdd("D", -1, DateAdd("M", 1, ("01/" + Format(DatePart("M", pdFecha), "00") + "/" + Format(DatePart("YYYY", pdFecha), "0000")))) = ldFecFin Then
                        SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCFF), 2))
                    Else
                        SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCF), 2))
                    End If
                End If
            End If
            numero = ""
        Else
            numero = numero + Mid(Ctamas, i, 1)
            If (i = Len(Ctamas)) Then
                Set oRep = New DRepFormula
                Set rsSaldo = oRep.ObtenerRep2SubColCtaSaldo(pdFecha, numero, pnMoneda)
                If pnMoneda = 1 Then
                    If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                        SaldoMas = SaldoMas + rsSaldo!nSaldo
                    End If
                Else
                    If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                       'Modificado PASIERS1382014
                        'SaldoMas = SaldoMas + Round(rsSaldo!nSaldo / lnTCFF, 2)
                        If (pdFecha = ldFecFin) And DateAdd("D", -1, DateAdd("M", 1, ("01/" + Format(DatePart("M", pdFecha), "00") + "/" + Format(DatePart("YYYY", pdFecha), "0000")))) = ldFecFin Then
                            SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCFF), 2))
                        Else
                            SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCF), 2))
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    For i = 1 To Len(Ctamenos)
        If (Mid(Ctamenos, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rsSaldo = oRep.ObtenerRep2SubColCtaSaldo(pdFecha, numero, pnMoneda)
            If pnMoneda = 1 Then
                If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                    SaldoMenos = SaldoMenos + rsSaldo!nSaldo
                End If
            Else
                If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                    'Modificado PASIERS1382014
                    'SaldoMas = SaldoMas + Round(rsSaldo!nSaldo / lnTCFF, 2)
                    If (pdFecha = ldFecFin) And DateAdd("D", -1, DateAdd("M", 1, ("01/" + Format(DatePart("M", pdFecha), "00") + "/" + Format(DatePart("YYYY", pdFecha), "0000")))) = ldFecFin Then
                        SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCFF), 2))
                    Else
                        SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCF), 2))
                    End If
                End If
            End If
            numero = ""
        Else
            numero = numero + Mid(Ctamenos, i, 1)
            If (i = Len(Ctamenos)) Then
                Set oRep = New DRepFormula
                Set rsSaldo = oRep.ObtenerRep2SubColCtaSaldo(pdFecha, numero, pnMoneda)
                If pnMoneda = 1 Then
                    If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                        SaldoMenos = SaldoMenos + rsSaldo!nSaldo
                    End If
                Else
                    If Not rsSaldo.EOF And Not rsSaldo.BOF Then
                        'Modificado PASIERS1382014
                        'SaldoMas = SaldoMas + Round(rsSaldo!nSaldo / lnTCFF, 2)
                        If (pdFecha = ldFecFin) And DateAdd("D", -1, DateAdd("M", 1, ("01/" + Format(DatePart("M", pdFecha), "00") + "/" + Format(DatePart("YYYY", pdFecha), "0000")))) = ldFecFin Then
                            SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCFF), 2))
                        Else
                            SaldoMas = SaldoMas + IIf(IsNull(rsSaldo!nSaldo), 0, Round((rsSaldo!nSaldo / lnTCF), 2))
                        End If
                    End If
                End If
            End If
        End If
    Next i
    Ctamas = ""
    Ctamenos = ""
    DevolverSaldoFormulaRep2 = (SaldoMas - SaldoMenos)
End Function
Private Sub CabeceraRep2()
    Set oRep = New DRepFormula
    Dim lx, lreg, lfil, lcol, lz, ly, i As Integer
    Dim ltemp As Boolean
    Dim lcolumna As String
    Dim oValoresCab As ADODB.Recordset
    Dim sCadena As String
    Set oValoresCab = New ADODB.Recordset
    Set oValoresCab = oRep.ObtenerCabeceraRep2FormulaBCR(CInt(lsMoneda))
    lx = 2
    ltemp = False
    lfil = 9
    lcol = 2
    lz = 0
    ly = 1
        
        xlHoja1.Cells(lfil, 1) = "Nombre Institución"
        xlHoja1.Cells(lfil + 1, 1) = "Cod. Swift"
        xlHoja1.Cells(lfil + 2, 1) = "Cod. Operación"
        
        xlHoja1.Cells(lfil + 4, 1) = "DIA"
        Do While Not oValoresCab.EOF
            If Not ltemp Then
                lreg = oValoresCab!cCodigo
                xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodSwif
                'i = oValoresCab!nValor
                'sCadena = oValoresCab!cCodigo
                xlHoja1.Cells(lfil + 2, lx) = oValoresCab!cCodigo 'IIf(i = 2, Mid(sCadena, 1, 2) + "1000", oValoresCab!cCodigo)
                ltemp = True
                lx = lx + 1
                ly = ly + 1
                lz = lz + 1
                lcolumna = oValoresCab!columna
            Else
                If (oValoresCab!cCodigo) = lreg Or (Mid(oValoresCab!cCodigo, 1, 2) = Mid(lreg, 1, 2)) Then
                    xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                    xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodSwif
                    'i = oValoresCab!nValor
                    'sCadena = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil + 2, lx) = oValoresCab!cCodigo 'IIf(i = 2, Mid(sCadena, 1, 2) + "1000", oValoresCab!cCodigo)
                    lx = lx + 1
                    lz = lz + 1
                    lcolumna = oValoresCab!columna
                Else
                    xlHoja1.Cells(lfil - 1, ly) = lcolumna
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                    xlHoja1.Cells(lfil, lx) = oValoresCab!cDescripcion
                    xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cCodSwif
                    'i = oValoresCab!nValor
                    'sCadena = oValoresCab!cCodigo
                    xlHoja1.Cells(lfil + 2, lx) = oValoresCab!cCodigo 'IIf(i = 2, Mid(sCadena, 1, 2) + "1000", oValoresCab!cCodigo)
                    lcolumna = oValoresCab!columna
                    ly = lx
                    lz = 1
                    lx = lx + 1
                    lreg = oValoresCab!cCodigo
                End If
            End If
            oValoresCab.MoveNext
        Loop
            RSClose oValoresCab
            xlHoja1.Cells(lfil - 1, ly) = lcolumna
            xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
            Set oRep = Nothing
End Sub
Private Function ImprimeObligacionSujetaEnc()
    CabeceraReporte
    ContenidoObligacionSujetaEnc
End Function
Private Sub CabeceraReporte()
    'PASI20140422
    Dim lsAnexo As String
    Select Case Right(lsCodOpe, 1)
        Case "6"
            lsAnexo = "1"
        Case "7"
            lsAnexo = "2"
        Case "8"
            lsAnexo = "3"
        Case "9"
            lsAnexo = "4"
    End Select
    ExcelAddHoja "ANX" & lsAnexo, xlLibro, xlHoja1, False
    'END PASI
    'ExcelAddHoja "Anx_" & Right(lsCodOpe, 1), xlLibro, xlHoja1, False 'comentado x pasi
    
    xlAplicacion.Range("A1:R100").Font.Size = 10
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 55
    
    If Mid(lsCodOpe, 6, 1) = "6" Then
        xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERÚ - DEPARTAMENTO DE ADMINISTRACIÓN DE ENCAJES"
        xlHoja1.Cells(4, 1) = IIf(lsMoneda = "1", "Reporte Nº 1 (EN NUEVOS SOLES)", "Reporte Nº 1 (EN US$.)")
    ElseIf Mid(lsCodOpe, 6, 1) = "7" Then
        xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERÚ - DEPARTAMENTO DE ADMINISTRACIÓN DE ENCAJES"
        xlHoja1.Cells(2, 1) = "OBLIGACIONES EXONERADAS DE GUARDAR ENCAJE CON INSTITUCIONES FINANCIERAS DEL PAIS"
        xlHoja1.Cells(4, 1) = IIf(lsMoneda = "1", "Reporte Nº 2 (EN NUEVOS SOLES)", "Reporte Nº 2 (EN US$.)")
    ElseIf Mid(lsCodOpe, 6, 1) = "8" Then
        xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERÚ - DEPARTAMENTO DE ADMINISTRACIÓN DE ENCAJES"
        xlHoja1.Cells(2, 1) = "OBLIGACIONES CON INSTITUCIONES FINANCIERAS DEL EXTERIOR"
        xlHoja1.Cells(4, 1) = IIf(lsMoneda = "1", "Reporte Nº 3 (EN NUEVOS SOLES)", "Reporte Nº 3 (EN US$.)")
    ElseIf Mid(lsCodOpe, 6, 1) = "9" Then
        xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERÚ - DEPARTAMENTO DE ADMINISTRACIÓN DE ENCAJES"
        xlHoja1.Cells(2, 1) = "OTRAS OBLIGACIONES NO SUJETAS A ENCAJE"
        xlHoja1.Cells(4, 1) = IIf(lsMoneda = "1", "Reporte Nº 4 (EN NUEVOS SOLES)", "Reporte Nº 4 (EN US$.)")
    End If
        xlHoja1.Cells(5, 1) = UCase("INSTITUCION : CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A.")
        xlHoja1.Cells(6, 1) = UCase("Periodo: " & EmitePeriodo)
    xlHoja1.Range("A1:A6").Font.Bold = True
    
    
End Sub
Private Sub ContenidoObligacionSujetaEnc()
    CabeceraRep1
    ContenidoSaldosRep1
End Sub
Private Sub CabeceraRep1()
    Set oRep = New DRepFormula
    Dim lx, lreg, lfil, lcol, lz, ly, i As Integer
    Dim ltemp As Boolean
    Dim lRegimen As String
    Dim oValoresCab As ADODB.Recordset
    Set oValoresCab = New ADODB.Recordset
    Set oValoresCab = oRep.ObtenerCabeceraRep1BaseFormulaBCR(CInt(lsMoneda))
    lx = 2
    ltemp = False
    lfil = 11
    lcol = 2
    lz = 0
    ly = 1
    
        xlHoja1.Cells(lfil, 1) = "Dia"
        xlHoja1.Range(xlHoja1.Cells(lfil - 1, 1), xlHoja1.Cells(lfil + 3, 1)).MergeCells = True
        
         Do While Not oValoresCab.EOF
            If Not ltemp Then
                lreg = oValoresCab!nCodRegimen
                xlHoja1.Cells(lfil, lx) = "(" & oValoresCab!cNumCol & ")"
                xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cDescripcion
                xlHoja1.Range(xlHoja1.Cells(lfil + 1, lx), xlHoja1.Cells(lfil + 3, lx)).MergeCells = True
                xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodigo
                ltemp = True
                lx = lx + 1
                ly = ly + 1
                lz = lz + 1
                lRegimen = oValoresCab!Regimen
            Else
                If oValoresCab!nCodRegimen = lreg Then
                    xlHoja1.Cells(lfil, lx) = "(" & oValoresCab!cNumCol & ")"
                    xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cDescripcion
                    xlHoja1.Range(xlHoja1.Cells(lfil + 1, lx), xlHoja1.Cells(lfil + 3, lx)).MergeCells = True
                    xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodigo
                    lx = lx + 1
                    lz = lz + 1
                    lRegimen = oValoresCab!Regimen
                Else
                    xlHoja1.Cells(lfil - 1, ly) = lRegimen
                    xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
                    xlHoja1.Cells(lfil, lx) = "(" & oValoresCab!cNumCol & ")"
                    xlHoja1.Cells(lfil + 1, lx) = oValoresCab!cDescripcion
                    xlHoja1.Range(xlHoja1.Cells(lfil + 1, lx), xlHoja1.Cells(lfil + 3, lx)).MergeCells = True
                    xlHoja1.Cells(lfil + 4, lx) = oValoresCab!cCodigo
                    lRegimen = oValoresCab!Regimen 'agregado pasi 20140304
                    ly = lx
                    lz = 1
                    lx = lx + 1
                    lreg = oValoresCab!nCodRegimen
                End If
            End If
            oValoresCab.MoveNext
        Loop
            RSClose oValoresCab
        xlHoja1.Cells(lfil - 1, ly) = lRegimen
        xlHoja1.Range(xlHoja1.Cells(lfil - 1, ly), xlHoja1.Cells(lfil - 1, ly + lz - 1)).Merge True
        xlHoja1.Cells(lfil - 2, 1) = "OBLIGACIONES SUJETAS A ENCAJE EN " & IIf(lsMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
        xlHoja1.Range(xlHoja1.Cells(lfil - 2, 1), xlHoja1.Cells(lfil - 2, lx - 1)).Merge True
        Set oRep = Nothing
End Sub
Private Sub ContenidoSaldosRep1()
    Set oRep = New DRepFormula
    Dim ldFecha As Date
    'Dim ldFechaAnt As String
    Dim rsRep1 As ADODB.Recordset
    Set rsRep1 = New ADODB.Recordset
    Dim lncol, lnfil, i, j As Integer
    Dim lnSaldo As Currency
    Dim lnTotalCol As Integer
    Dim nSaldoAnt As Currency
    Dim oEst As New NEstadisticas
    Dim SaldoCajAnt As Currency
    
    lNumColCAJAMERep1 = oRep.ObtieneRep1NumColCAJAME
    ldFecha = ldFecIni
    ldFechaAnt = DateAdd("d", -1, ldFecIni)
    SaldoCajAnt = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32")
    Set rsRep1 = oRep.ObtenerColRep1FormulaBCR(lsMoneda)
    lnTotalCol = rsRep1.RecordCount
    ReDim lsSaldos(1 To lnTotalCol, 0 To 0)
    lncol = 0
    lnfil = 0
    Do While ldFecha <= ldFecFin
        lnfil = lnfil + 1
        ReDim Preserve lsSaldos(1 To lnTotalCol, lnfil)
        Do While Not rsRep1.EOF
            lncol = lncol + 1
            If rsRep1!nValor = "1" Then
                lnSaldo = DevolverSaldoFormulaRep1(rsRep1!cvalor, ldFecha, rsRep1!cNumCol, rsRep1!cColDesc)
            End If
            If rsRep1!nValor = "2" Then
                lnSaldo = DevolverSaldoTotalizadoRep1(rsRep1!cvalor, ldFecha, lsMoneda)
            End If
            If rsRep1!nValor = "3" Then
                If (lsMoneda = "1") Then
                    lnSaldo = SaldoCajAnt
                End If
            End If
            lsSaldos(lncol, lnfil) = lnSaldo
            rsRep1.MoveNext
        Loop
        rsRep1.MoveFirst
        lncol = 0
        ldFecha = DateAdd("d", 1, ldFecha)
    Loop
    Set oRep = Nothing
    'Ingresar los datos al excel
    lnfil = 16
    lncol = 0
    
    Dim lnSaldoCelda() As Currency
    ReDim lnSaldoCelda(1 To UBound(lsSaldos, 1), 1)
    
    For i = 1 To UBound(lnSaldoCelda, 1)
        lnSaldoCelda(i, 1) = 0
    Next i
    
    For i = 1 To UBound(lsSaldos, 2)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, lncol) = i
        For j = 1 To UBound(lsSaldos, 1)
            lncol = lncol + 1
            xlHoja1.Cells(lnfil, lncol) = Format(lsSaldos(j, i), "#,#0.00")
            lnSaldoCelda(j, 1) = lnSaldoCelda(j, 1) + lsSaldos(j, i)
        Next j
        lncol = 0
        lnfil = lnfil + 1
    Next i
    lncol = 1
    For i = 1 To UBound(lnSaldoCelda, 1)
        lncol = lncol + 1
        xlHoja1.Cells(lnfil, 1) = "TOTAL"
        xlHoja1.Cells(lnfil, lncol) = Format(lnSaldoCelda(i, 1), "#,#0.00")
    Next i
End Sub
Private Function DevolverSaldoTotalizadoRep1(ByVal psValor As String, ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Set oRep = New DRepFormula
    Dim oEst As New NEstadisticas
    Dim SaldoMas, SaldoMenos, i As Currency
    Dim rRep1 As ADODB.Recordset
    Set rRep1 = New ADODB.Recordset
    Dim numero As String
    numero = ""
    SaldoMas = 0
    SaldoMenos = 0
    GeneraColumnasGenerales (psValor)
     For i = 1 To Len(colmas)
        If (Mid(colmas, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
            If (rRep1!nValor = 1) Then
                SaldoMas = SaldoMas + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASIERS1382014 agrego numero
            End If
            If (rRep1!nValor = 2) Then
                SaldoMas = SaldoMas + DevolverSaldoTotColRep1(numero, pdFecha, pnMoneda)
            End If
            If (rRep1!nValor = 3) Then
                If (pnMoneda = "1") Then
                    SaldoMas = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32")
                End If
            End If
            numero = ""
        Else
            numero = numero + Mid(colmas, i, 1)
            If (i = Len(colmas)) Then
                Set oRep = New DRepFormula
                Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
                If (rRep1!nValor = 1) Then
                    SaldoMas = SaldoMas + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASI20150310 agrego numero
                End If
                If (rRep1!nValor = 2) Then
                    SaldoMas = SaldoMas + DevolverSaldoTotColRep1(numero, pdFecha, pnMoneda)
                End If
                If (rRep1!nValor = 3) Then
                    If (pnMoneda = "1") Then
                        SaldoMas = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32")
                    End If
                End If
            End If
            'numero = ""
        End If
     Next i
     For i = 1 To Len(colmenos)
        If (Mid(colmas, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
            If (rRep1!nValor = 1) Then
                SaldoMenos = SaldoMenos + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha)
            End If
            If (rRep1!nValor = 2) Then
                SaldoMenos = SaldoMenos + DevolverSaldoTotColRep1(numero, pdFecha, pnMoneda)
            End If
            If (rRep1!nValor = 3) Then
                If (pnMoneda = "1") Then
                    SaldoMenos = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32")
                End If
            End If
            numero = ""
        Else
            numero = numero + Mid(colmenos, i, 1)
            If (i = Len(colmenos)) Then
                Set oRep = New DRepFormula
                Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
                If (rRep1!nValor = 1) Then
                    SaldoMenos = SaldoMenos + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha)
                End If
                If (rRep1!nValor = 2) Then
                    SaldoMenos = SaldoMenos + DevolverSaldoTotColRep1(numero, pdFecha, pnMoneda)
                End If
                If (rRep1!nValor = 3) Then
                    If (pnMoneda = "1") Then
                    SaldoMenos = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32")
                    End If
                End If
            End If
        End If
     Next i
     colmas = ""
     colmenos = ""
     DevolverSaldoTotalizadoRep1 = (SaldoMas - SaldoMenos)
End Function
Public Function DevolverSaldoTotColRep1(ByVal psValor As String, ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Set oRep = New DRepFormula
    Dim lSaldoMas, lSaldoMenos As Currency
    Dim lcolmas, lcolmenos As String
    Dim ltemp As Boolean
    Dim i As Integer
    Dim numero As String
    Dim lsValor As String
    Dim rRep1 As ADODB.Recordset
    Set rRep1 = New ADODB.Recordset
    numero = ""
    ltemp = True
    Set rRep1 = oRep.ObtenerRep1ValorColumna(CInt(psValor), pnMoneda)
    lsValor = rRep1!cvalor
    For i = 1 To Len(lsValor)
        If ltemp Then
            If Mid(lsValor, i, 1) = "+" Then
                lcolmas = lcolmas + ","
                ltemp = True
            Else
                If Mid(lsValor, i, 1) = "-" Then
                    ltemp = False
                    If Len(lcolmenos) <> 0 Then
                        lcolmenos = lcolmenos + ","
                    End If
                Else
                    lcolmas = lcolmas + Mid(lsValor, i, 1)
                    ltemp = True
                End If
            End If
        Else
            If Mid(lsValor, i, 1) = "-" Then
                lcolmenos = lcolmenos + ","
                ltemp = False
            Else
                If Mid(lsValor, i, 1) = "+" Then
                    ltemp = True
                    If Len(lcolmas) <> 0 Then
                        lcolmas = lcolmas + ","
                    End If
                Else
                    lcolmenos = lcolmenos + Mid(lsValor, i, 1)
                    ltemp = False
                End If
            End If
        End If
    Next i
    
    For i = 1 To Len(lcolmas)
        If (Mid(lcolmas, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
            lSaldoMas = lSaldoMas + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASIERS1382014 agrego numero,PASI20150310 agrego cColDesc
            numero = ""
        Else
            numero = numero + Mid(lcolmas, i, 1)
            If (i = Len(lcolmas)) Then
                Set oRep = New DRepFormula
                Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
                lSaldoMas = lSaldoMas + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASIERS1382014 agrego numero,PASI20150310 agrego cColDesc
                numero = ""
            End If
        End If
    Next i
    For i = 1 To Len(lcolmenos)
        If (Mid(lcolmenos, i, 1) = ",") Then
            Set oRep = New DRepFormula
            Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
            lSaldoMenos = lSaldoMenos + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASIERS1382014 agrego numero, PASI20150310 agrego cColDesc
            numero = ""
        Else
            numero = numero + Mid(lcolmenos, i, 1)
            If (i = Len(lcolmenos)) Then
                Set oRep = New DRepFormula
                Set rRep1 = oRep.ObtenerRep1ValorColumna(numero, pnMoneda)
                lSaldoMenos = lSaldoMenos + DevolverSaldoFormulaRep1(rRep1!cvalor, pdFecha, numero, rRep1!cColDesc) 'PASIERS1382014 agrego numero, PASI20150310 agrego cColDesc
                numero = ""
            End If
        End If
    Next i
    DevolverSaldoTotColRep1 = (lSaldoMas - lSaldoMenos)
End Function
Private Function DevolverSaldoFormulaRep1(ByVal psValor As String, ByVal pdFecha As Date, Optional ByVal nCol As Integer = 0, Optional ByVal cColDesc As String = "") As Currency 'PASIERS1382014 agrego NCOL
    Set oRep = New DRepFormula
    Dim lnSaldoCtas As Currency
    Dim rsSaldo As ADODB.Recordset
    Set rsSaldo = New ADODB.Recordset
    GeneraCuentas (psValor)
    Set rsSaldo = oRep.ObtenerRep1ColumnaCtaSaldo(Ctamas, Ctamenos, pdFecha)
    lnSaldoCtas = 0
    If Not RSVacio(rsSaldo) Then
        If Mid(lsCodOpe, 3, 1) = "1" Then
            lnSaldoCtas = IIf(IsNull(rsSaldo!Total), 0, rsSaldo!Total)
        Else
            If (UCase(cColDesc = "CAJA") Or nCol = lNumColCAJAMERep1) Then 'PASIERS1382014
                lnSaldoCtas = Round(IIf(IsNull(rsSaldo!TotalME), 0, rsSaldo!TotalME), 1)
            Else
                If (pdFecha = ldFecFin) And DateAdd("D", -1, DateAdd("M", 1, ("01/" + Format(DatePart("M", pdFecha), "00") + "/" + Format(DatePart("YYYY", pdFecha), "0000")))) = ldFecFin Then
                    lnSaldoCtas = IIf(IsNull(rsSaldo!Total), 0, Round((rsSaldo!Total / lnTCFF), 2))
                Else
                    lnSaldoCtas = IIf(IsNull(rsSaldo!Total), 0, Round((rsSaldo!Total / lnTCF), 2))
                End If
            End If
        End If
    End If
    Set oRep = Nothing
    Ctamas = ""
    Ctamenos = ""
    DevolverSaldoFormulaRep1 = lnSaldoCtas
    RSClose rsSaldo
End Function
Private Sub GeneraCuentas(ByVal psValor As String)
    Dim ltemp As Boolean
    Dim i As Integer
    ltemp = True
    For i = 1 To Len(psValor)
        If ltemp Then
            If Mid(psValor, i, 1) = "+" Then
                Ctamas = Ctamas + ","
                ltemp = True
            Else
                If Mid(psValor, i, 1) = "-" Then
                    ltemp = False
                    If Len(Ctamenos) <> 0 Then
                        Ctamenos = Ctamenos + ","
                    End If
                Else
                    Ctamas = Ctamas + Mid(psValor, i, 1)
                    ltemp = True
                End If
            End If
        Else
            If Mid(psValor, i, 1) = "-" Then
                Ctamenos = Ctamenos + ","
                ltemp = False
            Else
                If Mid(psValor, i, 1) = "+" Then
                    ltemp = True
                    If Len(Ctamas) <> 0 Then
                        Ctamas = Ctamas + ","
                    End If
                Else
                    Ctamenos = Ctamenos + Mid(psValor, i, 1)
                    ltemp = False
                End If
            End If
        End If
    Next
End Sub
Private Sub GeneraColumnasGenerales(ByVal psValor As String)
    Dim ltemp As Boolean
    Dim i As Integer
    ltemp = True
    For i = 1 To Len(psValor)
        If ltemp Then
            If Mid(psValor, i, 1) = "+" Then
                colmas = colmas + ","
                ltemp = True
            Else
                If Mid(psValor, i, 1) = "-" Then
                    ltemp = False
                    If Len(colmenos) <> 0 Then
                        colmenos = colmenos + ","
                    End If
                Else
                    colmas = colmas + Mid(psValor, i, 1)
                    ltemp = True
                End If
            End If
        Else
            If Mid(psValor, i, 1) = "-" Then
                colmenos = colmenos + ","
                ltemp = False
            Else
                If Mid(psValor, i, 1) = "+" Then
                    ltemp = True
                    If Len(colmas) <> 0 Then
                        colmas = colmas + ","
                    End If
                Else
                    colmenos = colmenos + Mid(psValor, i, 1)
                    ltemp = False
                End If
            End If
        End If
    Next i
End Sub
Private Function EmitePeriodo() As String
    Dim lsPeriodo As String
    lsPeriodo = "#" & FillNum(Trim(Str(Month(ldFecIni))), 2, "0")
    lsPeriodo = lsPeriodo & " (Del " & ldFecIni & " AL " & ldFecFin & " : Mes de " & Format(ldFecIni, "mmmm yyyy") & ")"
    EmitePeriodo = lsPeriodo
End Function

Private Sub Form_Load()
    Set oRep = New DRepFormula
End Sub

