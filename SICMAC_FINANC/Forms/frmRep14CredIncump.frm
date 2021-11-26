VERSION 5.00
Begin VB.Form frmRep14CredIncump 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte 14: Reporte de Créditos Según dias de Incumplimiento"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRep14CredIncump"
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
 
Dim sservidorconsolidada As String
     
Public Sub ImprimeReporte14(psOpeCod As String, pdFecha As Date, nTipCambio As Double)
On Error GoTo GeneraEstadError
Dim oConecta As DConecta
Dim sSql As String
Dim rs   As ADODB.Recordset



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
   lsArchivo = App.path & "\SPOOLER\" & "SBSRep14_" & Format(pdFecha, "mmyyyy") & ".xlsx" 'PASIERS1332014 se ha cambiado la extension de xls a xlsx
   
   
   
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbExcel Then
      ExcelAddHoja "Rep_14", xlLibro, xlHoja1
      Genera14 pdFecha, nTipCambio
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

Private Sub Genera14(pdFecha As Date, pnTipoCambio As Double)
    Dim lbExisteHoja  As Boolean
    Dim i  As Integer
    Dim j As Integer
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim lsMon As String
    
    Dim sSql As String
    Dim sSql2 As String '***Agregado por PASI20140102 TI-ERS161-2013
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset '***Agregado por PASI20140102 TI-ERS161-2013
    Dim sCadenaCuentas As String
     
    Dim mMN As Integer
    Dim mME As Integer
    
    Dim cVigente As String
    Dim cPigno As String
    
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "','2201','2205'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
    
    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmReportes
    oBarra.Max = 2
    oBarra.Progress 0, "REPORTE 14: CRED DIAS INCUMPLIMIENTO", "Cargando Datos", "", vbBlue
    
    lnFila = CabeceraReporte(pdFecha, "CREDITOS SEGUN DIAS DE INCUMPLIMIENTO", "REPORTE 14", 21)
    
    If gbBitCentral = True Then
        sCadenaCuentas = sCadenaCuentas & " cProducto = SubString(CCT.cTpoCredCod, 1,1),"
        sCadenaCuentas = sCadenaCuentas & " cRango = case when CC.ndiasatraso<=0 then '01. Sin Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 1 And 15 then '02. Entre 01 y 15 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 16 And 30 then '03. Entre 16 y 30 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 31 And 60 then '04. Entre 31 y 60 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 61 And 90 then '05. Entre 61 y 90 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 91 And 120 then '06. Entre 91 y 120 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 121 And 180 then '07. Entre 121 y 180 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 181 And 365 then '08. Entre 181 y 365 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " When CC.ndiasatraso >365 then '09. Mayor a 365 dias de atraso' end "

        sSql = "Select "
        sSql = sSql & sCadenaCuentas & ", "
        sSql = sSql & " nTotal = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then cc.nsaldocap else cc.nsaldocap * " & pnTipoCambio & " end),  "
        sSql = sSql & " nporcionnoamortizada = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then cc.ncapvencido else cc.ncapvencido * " & pnTipoCambio & " end), "
        sSql = sSql & " nSaldo = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then (cc.nSaldoCap - cc.nCapVencido) else (cc.nSaldoCap - cc.nCapVencido)* " & pnTipoCambio & " end) "
        sSql = sSql & " From " & sservidorconsolidada & "creditoconsol cc "
        sSql = sSql & " inner join " & sservidorconsolidada & "creditoconsoltotal cct on cct.cCtaCod = cc.cCtaCod"
        sSql = sSql & " Where "
        sSql = sSql & " Cc.nPrdEstado in (" & cVigente & ", " & cPigno & ") "
        sSql = sSql & " and cct.cTpoProdCod not in ('515','516') " '***Agregado por PASI20140102 TI-ERS161-2013
        
        sSql = sSql & " group by " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")
        sSql = sSql & " Order By " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")
        
        '***Agregado por PASI20140102 TI-ERS161-2013
        sSql2 = "Select "
        sSql2 = sSql2 & sCadenaCuentas & ", "
        sSql2 = sSql2 & " nTotal = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then cc.nsaldocap else cc.nsaldocap * " & pnTipoCambio & " end),  "
        sSql2 = sSql2 & " nporcionnoamortizada = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then cc.ncapvencido else cc.ncapvencido * " & pnTipoCambio & " end), "
        sSql2 = sSql2 & " nSaldo = Sum(case when substring(cc.cCtaCod, 9,1) = '1' then (cc.nSaldoCap - cc.nCapVencido) else (cc.nSaldoCap - cc.nCapVencido)* " & pnTipoCambio & " end) "
        sSql2 = sSql2 & " From " & sservidorconsolidada & "creditoconsol cc "
        sSql2 = sSql2 & " inner join " & sservidorconsolidada & "creditoconsoltotal cct on cct.cCtaCod = cc.cCtaCod"
        sSql2 = sSql2 & " Where "
        sSql2 = sSql2 & " Cc.nPrdEstado in (" & cVigente & ", " & cPigno & ") "
        sSql2 = sSql2 & " and cct.cTpoProdCod in ('515','516') "

        sSql2 = sSql2 & " group by " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")
        sSql2 = sSql2 & " Order By " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")
        '***Fin pasi
        
    Else
 
        sCadenaCuentas = "cMoneda = substring(cc.ccodcta, 6,1), cProducto = case when substring(CC.cCodCta, 3,1) in ('1') then '01. Comerciales'"
        sCadenaCuentas = sCadenaCuentas & " when substring(CC.cCodCta, 3,1) in ('2') then '02. MicroEmpresa' "
        sCadenaCuentas = sCadenaCuentas & " when substring(CC.cCodCta, 3,1) in ('3') then '03. Consumo' "
        sCadenaCuentas = sCadenaCuentas & " when substring(CC.cCodCta, 3,1) in ('4') then '04. Hipotecarios' end, "
        sCadenaCuentas = sCadenaCuentas & " cRango = case when CC.ndiasatraso<=0 then '01. Sin Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 1 And 15 then '02. Entre 01 y 15 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 16 And 30 then '03. Entre 16 y 30 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 31 And 60 then '04. Entre 31 y 60 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 61 And 90 then '05. Entre 61 y 90 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 91 And 120 then '06. Entre 91 y 120 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 121 And 180 then '07. Entre 121 y 180 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " when CC.ndiasatraso between 181 And 365 then '08. Entre 181 y 365 Dias de Atraso' "
        sCadenaCuentas = sCadenaCuentas & " When CC.ndiasatraso >365 then '09. Mayor a 365 dias de atraso' end "

        sSql = "Select "
        sSql = sSql & sCadenaCuentas & ", "
        sSql = sSql & " Sum(cc.nsaldocap) as nTotal, Sum(cc.ncapvencido) as nporcionnoamortizada, "
        sSql = sSql & " Sum(cc.nSaldoCap - cc.nCapVencido) as nSaldo "
        sSql = sSql & " From creditoconsol cc "
        sSql = sSql & " Where "
        sSql = sSql & " (Cc.cEstado in('F', '1','4','6','7') "
        sSql = sSql & " or (cc.cestado= 'V' AND cc.cCondCre = 'J')) "
        
        sSql = sSql & " group by " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")
        sSql = sSql & " Order By " & Replace(Replace(Replace(sCadenaCuentas, "cMoneda = ", ""), "cProducto = ", ""), "cRango = ", "")

    End If
    
    xlHoja1.Range("B13") = "Corporativos"
    xlHoja1.Range("B15") = "Tratados como corporativos"
    xlHoja1.Range("B17") = "Grandes empresas"
    xlHoja1.Range("B19") = "Medianas empresas"
    xlHoja1.Range("B21") = "Pequeñas empresas"
    xlHoja1.Range("B23") = "Micro empresas"
    xlHoja1.Range("B25") = "Consumo"
    xlHoja1.Range("B28") = "TOTAL"
    
    xlHoja1.Range("B13:B14").MergeCells = True
    xlHoja1.Range("B15:B16").MergeCells = True
    xlHoja1.Range("B17:B18").MergeCells = True
    xlHoja1.Range("B19:B20").MergeCells = True
    xlHoja1.Range("B21:B22").MergeCells = True
    xlHoja1.Range("B23:B24").MergeCells = True
    xlHoja1.Range("B25:B26").MergeCells = True
    xlHoja1.Range("B13:B26").WrapText = True
    
    xlHoja1.Range("C13") = "Arrend.Financiero + Capit.Inmobiliaria 4/"
    xlHoja1.Range("C14") = "Otros créditos corporativos 5/"
    xlHoja1.Range("C15") = "Arrend.Financiero + Capit.Inmobiliaria 6/"
    xlHoja1.Range("C16") = "Otros créditos tratados como corporativos 7/"
    xlHoja1.Range("C17") = "Arrend.Financiero + Capit.Inmobiliaria 8/"
    xlHoja1.Range("C18") = "Otros créditos a grandes empresas 9/"
    xlHoja1.Range("C19") = "Arrend.Financiero + Capit.Inmobiliaria 10/"
    xlHoja1.Range("C20") = "Otros créditos a medianas empresas 11/"
    xlHoja1.Range("C21") = "Arrend.Financiero + Capit.Inmobiliaria 12/"
    xlHoja1.Range("C22") = "Otros créditos a pequeñas empresas 13/"
    xlHoja1.Range("C23") = "Arrend.Financiero + Capit.Inmobiliaria 14/"
    xlHoja1.Range("C24") = "Otros créditos a micro empresas 15/"
    xlHoja1.Range("C25") = "Tarjeta de crédito 16/"
    xlHoja1.Range("C26") = "Otros créditos de consumo 17/"
    xlHoja1.Range("B27") = "Hipotecarios para vivienda 18/"
    xlHoja1.Range("B27:C27").MergeCells = True
    xlHoja1.Range("D13:T27") = "0"
    xlHoja1.Range("D13:T27").NumberFormat = "#,###,##0.00"
    xlHoja1.Range("C13").ColumnWidth = 40
    
    lsProd = "": lnFilaIni = lnFila: lsMon = ""
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.BOF Then
        oBarra.Progress 2, "REPORTE 14: CRED DIAS INCUMPLIMIENTO", "Cargando Datos", "", vbBlue
        oBarra.Max = rs.RecordCount
        Do While Not rs.EOF
            If rs!cProducto = 1 Then
                lnFila = rs!cProducto * 2 + 12
            ElseIf rs!cProducto = 6 Or rs!cProducto = 7 Or rs!cProducto = 8 Then
                lnFila = rs!cProducto + 19
            ElseIf rs!cProducto = 2 Or rs!cProducto = 3 Or rs!cProducto = 4 Or rs!cProducto = 5 Then
                lnFila = rs!cProducto * 2 + 14
            End If
            
            If Mid(rs!cRango, 1, 2) = "01" Then
                xlHoja1.Cells(lnFila, 4) = Format(rs!nTotal, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "02" Then
                xlHoja1.Cells(lnFila, 5) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 6) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "03" Then
                xlHoja1.Cells(lnFila, 7) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 8) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "04" Then
                xlHoja1.Cells(lnFila, 9) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 10) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "05" Then
                xlHoja1.Cells(lnFila, 11) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 12) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "06" Then
                xlHoja1.Cells(lnFila, 13) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 14) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "07" Then
                xlHoja1.Cells(lnFila, 15) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 16) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "08" Then
                xlHoja1.Cells(lnFila, 17) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 18) = Format(rs!nSaldo, "#,###.00")
            ElseIf Mid(rs!cRango, 1, 2) = "09" Then
                xlHoja1.Cells(lnFila, 19) = Format(rs!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 20) = Format(rs!nSaldo, "#,###.00")
            End If
            oBarra.Progress rs.Bookmark, "REPORTE 14: CRED DIAS INCUMPLIMIENTO", "Generando Reporte", "", vbBlue
            rs.MoveNext
        Loop
    End If
    
    '***Agregado por PASI20140102 TI-ERS161-2013
    Set rs2 = oCon.CargaRecordSet(sSql2)
    If Not rs2.BOF Then
        Do While Not rs2.EOF
            If rs2!cProducto = 1 Then
                lnFila = (rs2!cProducto * 2 + 12) - 1
            ElseIf rs2!cProducto = 6 Or rs2!cProducto = 7 Or rs2!cProducto = 8 Then
                lnFila = (rs2!cProducto + 19) - 1
            ElseIf rs2!cProducto = 2 Or rs2!cProducto = 3 Or rs2!cProducto = 4 Or rs2!cProducto = 5 Then
                lnFila = (rs2!cProducto * 2 + 14) - 1
            End If

            If Mid(rs2!cRango, 1, 2) = "01" Then
                xlHoja1.Cells(lnFila, 4) = Format(rs2!nTotal, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "02" Then
                xlHoja1.Cells(lnFila, 5) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 6) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "03" Then
                xlHoja1.Cells(lnFila, 7) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 8) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "04" Then
                xlHoja1.Cells(lnFila, 9) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 10) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "05" Then
                xlHoja1.Cells(lnFila, 11) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 12) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "06" Then
                xlHoja1.Cells(lnFila, 13) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 14) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "07" Then
                xlHoja1.Cells(lnFila, 15) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 16) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "08" Then
                xlHoja1.Cells(lnFila, 17) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 18) = Format(rs2!nSaldo, "#,###.00")
            ElseIf Mid(rs2!cRango, 1, 2) = "09" Then
                xlHoja1.Cells(lnFila, 19) = Format(rs2!nporcionnoamortizada, "#,###.00")
                xlHoja1.Cells(lnFila, 20) = Format(rs2!nSaldo, "#,###.00")
            End If
            rs2.MoveNext
        Loop
    End If
     '***Fin pasi
    
    For i = 13 To 27
        xlHoja1.Range("U" & i) = "= SUM(D" & i & ":T" & i & ")"
    Next
    
    xlHoja1.Range("D28") = "= SUM(D13:D27)"
    xlHoja1.Range("E28") = "= SUM(E13:E27)"
    xlHoja1.Range("F28") = "= SUM(F13:F27)"
    xlHoja1.Range("G28") = "= SUM(G13:G27)"
    xlHoja1.Range("H28") = "= SUM(H13:H27)"
    xlHoja1.Range("I28") = "= SUM(I13:I27)"
    xlHoja1.Range("J28") = "= SUM(J13:J27)"
    xlHoja1.Range("K28") = "= SUM(K13:K27)"
    xlHoja1.Range("L28") = "= SUM(L13:L27)"
    xlHoja1.Range("M28") = "= SUM(M13:M27)"
    xlHoja1.Range("N28") = "= SUM(N13:N27)"
    xlHoja1.Range("O28") = "= SUM(O13:O27)"
    xlHoja1.Range("P28") = "= SUM(P13:P27)"
    xlHoja1.Range("Q28") = "= SUM(Q13:Q27)"
    xlHoja1.Range("R28") = "= SUM(R13:R27)"
    xlHoja1.Range("S28") = "= SUM(S13:S27)"
    xlHoja1.Range("T28") = "= SUM(T13:T27)"
    xlHoja1.Range("U28") = "= SUM(U13:U27)"
    
    xlHoja1.Range("B13:U28").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("B10:C27").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("D13:U27").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("B28:U28").BorderAround xlContinuous, xlMedium
    
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    
    xlHoja1.Cells(30, 2) = "'1/"
    xlHoja1.Cells(33, 2) = "'2/"
    xlHoja1.Cells(34, 2) = "'3/"
    xlHoja1.Cells(35, 2) = "'4/"
    xlHoja1.Cells(37, 2) = "'5/"
    xlHoja1.Cells(39, 2) = "'6/"
    xlHoja1.Cells(41, 2) = "'7/"
    xlHoja1.Cells(42, 2) = "'8/"
    xlHoja1.Cells(43, 2) = "'9/"

    xlHoja1.Cells(30, 3) = "Se debe registrar el saldo total de créditos que presentan atrasos. Para definir el tramo considerar desde el primer día de incumplimiento del pago según la fecha contractual.  "
    xlHoja1.Cells(31, 3) = "Registrar la parte atrasada y el saldo correspondiente que aún no ha vencido en el tramo de mayor días de incumplimiento. "
    xlHoja1.Cells(32, 3) = "Para cada tramo, en porción no amortizada considerar el monto impago acumulado; y en saldo reportar la parte del crédito que aún no ha vencido conforme al cronograma de pagos establecido en el contrato."
    xlHoja1.Cells(33, 3) = "Corresponde a la suma de todas las columnas."
    xlHoja1.Cells(34, 3) = "Incluye operaciones VAC."
    xlHoja1.Cells(35, 3) = "Cuentas 140101+140301+140401+140501+140601 - 14010104-14010111-14010112  - 14030104-14030111-14030112 - 14040104-14040111-14040112 - 14050104-14050111-14050112 - 14060104-14060111-14060112. No incluye sobregiros ni operaciones de arrendamiento financiero;"
    xlHoja1.Cells(36, 3) = "los cuales se registran en fila aparte."
    xlHoja1.Cells(37, 3) = "Cuentas 140102+140302+140402+140502+140602 - 14010204-14010211-14010212  - 14030204-14030211 - 1403.0212-14040204-14040211-14040212 - 14050204-14050211-14050212 - 14060204-14060211-14060212. No incluye sobregiros ni operaciones de"
    xlHoja1.Cells(38, 3) = "arrendamiento financiero; los cuales se registran en fila aparte."
    xlHoja1.Cells(39, 3) = "Cuentas 140103+140403+140503+140603 - 14010304-14010311-14010312 - 14040304-14040311-14040312 - 14050304-14050311-14050312 - 14060304-14060311-14060312. No incluye sobregiros ni operaciones de arrendamiento financiero; los cuales se registran en "
    xlHoja1.Cells(40, 3) = "fila aparte."
    xlHoja1.Cells(41, 3) = "Cuentas 140104 + 140404 + 140504 + 140604."
    xlHoja1.Cells(42, 3) = "Cuentas 14010104 + 14010204 + 14010304 + 14030104+14030204+14040104+14040204+14040304+14050104 + 14050204 + 14050304 + 14060104 + 14060204 + 14060304. Considerar en el tramo que corresponde al número de días de otorgado el sobregiro."
    xlHoja1.Cells(43, 3) = "Cuentas 14010111+ 14010112 + 14030111+ 14030112+14040111 + 14040112 + 14050111+ 14050112 + 14060111 + 14060112 + 14010211+ 14010212 + 14030211+14030212 + 14040211+ 14040212 + 14050211+ 14050212 + 14060211 + 14060212 + 14010311 + 14010312 + 14040311+"
    xlHoja1.Cells(44, 3) = "14040312 + 14050311+ 14050312 + 14060311+ 14060312."

    xlHoja1.Cells(50, 3) = "GERENTE GENERAL"
    xlHoja1.Cells(50, 6) = "CONTADOR GENERAL"
    xlHoja1.Cells(51, 6) = "MATRICULA No"
    xlHoja1.Cells(50, 13) = "FUNCIONARIO RESPONSABLE"
    
    xlHoja1.Range("F49:H49").MergeCells = True
    xlHoja1.Range("F50:H50").MergeCells = True
    xlHoja1.Range("F51:H51").MergeCells = True
    
    xlHoja1.Range("M49:O49").MergeCells = True
    xlHoja1.Range("M50:O50").MergeCells = True
    
    xlHoja1.Range("B32:C46").Font.Size = 6
     
    xlHoja1.Range("C50:M51").HorizontalAlignment = xlCenter
     
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub

Private Function CabeceraReporte(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer) As Integer
Dim lnFila As Integer
Dim i As Integer

    xlHoja1.Range("A1:U100").Font.Size = 8

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
    
    lnFila = 1
    xlHoja1.Cells(lnFila, 3) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 20) = "REPORTE 14"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "EMPRESA : " & gsNomCmac
    
    xlHoja1.Range("C1:D1").MergeCells = True
    xlHoja1.Range("C3:D3").MergeCells = True
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 2) = psTitulo
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).HorizontalAlignment = xlHAlignCenter
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 2) = "Al " & Format(pdFecha, "dd mmmm yyyy")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).HorizontalAlignment = xlHAlignCenter
    
    lnFila = lnFila + 1
    '''xlHoja1.Cells(lnFila, 2) = "En Nuevos Soles" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila, 2) = "En " & StrConv(gcPEN_PLURAL, vbProperCase) & "" 'MARG ERS044-2016
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, pnCols)).HorizontalAlignment = xlHAlignCenter
    
    lnFila = lnFila + 2
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 2, 3)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila + 2, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 20)).MergeCells = True
     
    xlHoja1.Range(xlHoja1.Cells(lnFila, 21), xlHoja1.Cells(lnFila + 2, 21)).MergeCells = True
    
    lnFila = lnFila + 1

    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 12)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 14)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 15), xlHoja1.Cells(lnFila, 16)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 17), xlHoja1.Cells(lnFila, 18)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 19), xlHoja1.Cells(lnFila, 20)).MergeCells = True
    
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila + 1, pnCols)).Font.Bold = True
    
    xlHoja1.Cells(10, 4) = "Saldo de Creditos sin Atraso"
    xlHoja1.Cells(10, 5) = "Incumplimiento 1/"
    xlHoja1.Cells(11, 5) = "De 1 a 15 dias"
    xlHoja1.Cells(11, 7) = "Entre 16 y 30 dias"
    xlHoja1.Cells(11, 9) = "Entre 31 y 60 dias"
    xlHoja1.Cells(11, 11) = "Entre 61 y 90 dias"
    xlHoja1.Cells(11, 13) = "Entre 91 y 120 dias"
    xlHoja1.Cells(11, 15) = "Entre 121 y 180 dias"
    xlHoja1.Cells(11, 17) = "Entre 181 y 365 dias"
    xlHoja1.Cells(11, 19) = "Mayor a 365 dias"
    xlHoja1.Cells(10, 21) = "Saldo Total de Creditos 2/"
        
    For i = 5 To 19 Step 2
        xlHoja1.Cells(12, i) = "Porción no Amortizada"
        xlHoja1.Cells(12, i + 1) = "Saldo"
    Next
    
    CabeceraReporte = lnFila
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



