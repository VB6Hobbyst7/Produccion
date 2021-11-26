Attribute VB_Name = "ModuloReporteRRHH"
Option Explicit

Public Function CargaRAuditoria(Periodo As String) As ADODB.Recordset
   On Error GoTo CargaConsultaErr
   Dim psSql As String, dbConec As New DConecta
'    psSql = "sp_RRHHReporteAuditoria '" & Periodo & "'"
    psSql = "declare @Periodo CHAR(6) set @Periodo='" & Periodo & "'" & vbCrLf
    psSql = psSql & " SELECT     a.cRHCod,  w.cPersNombre,ISNULL(b.nMonto, 0) AS Sueldo , ISNULL(c.nMonto, 0) AS vac, ISNULL(d.nMonto, 0) AS Reintegro , ISNULL(e.nMonto, 0) AS SUBENFER, ISNULL(f.nMonto, 0) AS SUBPOST, " & _
            "ISNULL(g.nMonto, 0) AS Prod, ISNULL(h.nMonto, 0) AS Bono, ISNULL(i.nMonto, 0) AS Util, ISNULL(j.nMonto, 0) AS CTS, ISNULL(k.nMonto, 0) AS GRATIF, ISNULL(l.nMonto, 0) AS Aguinaldo , ISNULL(m.nMonto, 0)AS Otros, " & _
            "ISNULL(b.nMonto, 0) + ISNULL(c.nMonto, 0) + ISNULL(d.nMonto, 0) + ISNULL(e.nMonto, 0) + ISNULL(f.nMonto, 0) + ISNULL(g.nMonto, 0) + ISNULL(h.nMonto, 0) + ISNULL(i.nMonto, 0) + ISNULL(j.nMonto, 0) + ISNULL(k.nMonto, 0) + " & _
            "ISNULL(l.nMonto, 0) + ISNULL(m.nMonto, 0) AS Total FROM dbo.RRHH a Left OUTER JOIN " & _
            " (SELECT      cperscod, sum(nMonto)nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E01' AND crhconceptocod = '130' group by cperscod) b ON a.cPersCod = b.cperscod LEFT  outer JOIN " & vbCrLf
    psSql = psSql & " (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E06' AND crhconceptocod = '130'group by cperscod) c ON a.cPersCod = c.cperscod LEFT  outer JOIN(SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E13' AND crhconceptocod = '130' group by cperscod) d ON a.cPersCod = d.cperscod  LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E12' AND crhconceptocod = '130' group by cperscod) e ON a.cPersCod = e.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E07' AND crhconceptocod = '130' group by cperscod ) f ON a.cPersCod = f.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod = 'E15' AND crhconceptocod = '130 'group by cperscod) g ON a.cPersCod = g.cperscod LEFT OUTER JOIN " & vbCrLf
    psSql = psSql & " (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo LIKE @Periodo + '%' AND " & _
            "cPlanillaCod = 'E16' AND crhconceptocod = '130' group by cperscod ) h ON a.cPersCod = h.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo LIKE @Periodo + '%' AND " & _
            "cPlanillaCod = 'E04' AND crhconceptocod = '130' group by cperscod ) i ON a.cPersCod = i.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo LIKE @Periodo + '%' AND " & vbCrLf & _
            "cPlanillaCod = 'E05' AND crhconceptocod = '130' group by cperscod ) j ON a.cPersCod = j.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo LIKE @Periodo + '%' AND " & _
            "cPlanillaCod = 'E02' AND crhconceptocod = '130' group by cperscod ) k ON a.cPersCod = k.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(Isnull(nMonto,0))nMonto FROM rhplanilladetcon WHERE cRRHHPeriodo " & vbCrLf
    psSql = psSql & " LIKE @Periodo + '%' AND cPlanillaCod = 'E11' AND crhconceptocod = '130' group by cperscod ) l ON a.cPersCod = l.cperscod LEFT OUTER JOIN (SELECT     cperscod, sum(nMonto)nMonto FROM rhplanilladetcon " & _
            "WHERE cRRHHPeriodo LIKE @Periodo + '%' AND cPlanillaCod NOT IN ('E01', 'E06', 'E13', 'E12', 'E07', 'E15', 'E06', 'E04', 'E05', 'E02','E11') AND cPlanillaCod LIKE 'E%' AND crhconceptocod = '130' group by cperscod) m ON a.cPersCod = m.cperscod " & _
            "INNER JOIN dbo.Persona w ON a.cPersCod = w.cPersCod WHERE     (a.cRHCod LIKE 'E%') AND (ISNULL(b.nMonto, 0) + ISNULL(c.nMonto, 0) + ISNULL(d.nMonto, 0) + ISNULL(e.nMonto, 0) + ISNULL(f.nMonto, 0) + ISNULL(g.nMonto, 0) + " & _
            "ISNULL(h.nMonto, 0) + ISNULL(i.nMonto, 0) + ISNULL(j.nMonto, 0) + ISNULL(k.nMonto, 0) + ISNULL(l.nMonto, 0) + ISNULL(m.nMonto, 0) <> 0) ORDER BY a.cRHCod"
    dbConec.AbreConexion
    Set CargaRAuditoria = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   dbConec.CierraConexion
   Exit Function
   MsgBox dbConec.CadenaConexion
CargaConsultaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CargaBalanceGeneral Method")
End Function

Public Function CargaConsultaPlanillas(Periodo As String) As ADODB.Recordset
   On Error GoTo CargaConsultaErr
   Dim psSql As String, dbConec As New DConecta
    psSql = "select cRHCod,cPersNombre, Veces, monto from RRHH R " & _
            "Inner Join (select cPerscod, count(*) Veces, sum(nmonto) Monto from RHPlanillaDetCon P where cPlanillaCod = 'E15' and cRHConceptoCod ='130' and cRRHHPeriodo like '" & Periodo & "%' group by cPerscod) X on X.cPersCod = R.cPersCod " & _
            "Inner Join Persona P on P.cPersCod = x.cPerscod Order by cRHCod"
    dbConec.AbreConexion
    Set CargaConsultaPlanillas = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   dbConec.CierraConexion
   Exit Function
   MsgBox dbConec.CadenaConexion
CargaConsultaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CargaBalanceGeneral Method")
End Function

Public Sub Generar_ReporteAuditoria(Periodo As String)

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    Dim liLineaInicio As Integer
    Dim liLineas As Integer
    Dim liLineasTemp As Integer
    Dim licont As Integer
    Dim liNro As Integer

    Dim i As Integer
    Dim lnTotal As Double

    Dim glsArchivo As String

    Dim CadTit As String
    Dim CadImp   As String
    Dim lnDivide As Integer
    Dim RsRRHH As New ADODB.Recordset
    
    Dim sTit1 As String
    Dim sTit As String
    Dim sTit2 As String
    Dim sTit3 As String
    Dim sTit4 As String
    Dim nPos As Integer
    Dim Fila As Integer
    Dim Año As Integer
    Dim mes As Integer
    

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Item(1)
    
    Dim lsNomHoja As String
        

        '**************************************************************************
        Año = Right(Periodo, 4)
        mes = Mid(Periodo, 4, 2)
        xlHoja1.Name = Año & Right("00" & mes, 2)
        
'        MsgBox Año & Right("00" & mes, 2)

         Set RsRRHH = CargaRAuditoria(Año & Right("00" & mes, 2))
         
'         Set dBalance = Nothing

         xlAplicacion.Range("A1:A1").ColumnWidth = 10
         xlAplicacion.Range("B1:B1").ColumnWidth = 40
         xlAplicacion.Range("C1:C1").ColumnWidth = 10
         xlAplicacion.Range("D1:D1").ColumnWidth = 10
         xlAplicacion.Range("E1:E1").ColumnWidth = 10
         xlAplicacion.Range("F1:F1").ColumnWidth = 10
         xlAplicacion.Range("G1:G1").ColumnWidth = 10
         xlAplicacion.Range("H1:H1").ColumnWidth = 10
         xlAplicacion.Range("I1:I1").ColumnWidth = 10
         xlAplicacion.Range("J1:J1").ColumnWidth = 10
         xlAplicacion.Range("K1:K1").ColumnWidth = 10
         xlAplicacion.Range("L1:L1").ColumnWidth = 10
         xlAplicacion.Range("M1:M1").ColumnWidth = 10
         xlAplicacion.Range("N1:N1").ColumnWidth = 10
         xlAplicacion.Range("O1:O1").ColumnWidth = 10

'         CadTit = Trim(prs!cDescrip)

         xlHoja1.Cells(1, 1) = gsNomCmac
         xlHoja1.Cells(1, 14) = "Fecha: " & Date
         xlHoja1.Cells(3, 1) = "R E P O R T E    P A R A    A U D I T O R I A"

         xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 14)).Font.Bold = True
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 15)).Merge True
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 15)).HorizontalAlignment = xlCenter
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 15)).Font.Bold = True
         xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 15)).Font.Bold = True

         liLineas = 6
        If Not RsRRHH.EOF Then
            liLineas = liLineas + 1
            xlHoja1.Cells(liLineas, 1) = RsRRHH(0).Name
            xlHoja1.Cells(liLineas, 2) = RsRRHH(1).Name
            xlHoja1.Cells(liLineas, 3) = RsRRHH(2).Name
            xlHoja1.Cells(liLineas, 4) = RsRRHH(3).Name
            xlHoja1.Cells(liLineas, 5) = RsRRHH(4).Name
            xlHoja1.Cells(liLineas, 6) = RsRRHH(5).Name
            xlHoja1.Cells(liLineas, 7) = RsRRHH(6).Name
            xlHoja1.Cells(liLineas, 8) = RsRRHH(7).Name
            xlHoja1.Cells(liLineas, 9) = RsRRHH(8).Name
            xlHoja1.Cells(liLineas, 10) = RsRRHH(9).Name
            xlHoja1.Cells(liLineas, 11) = RsRRHH(10).Name
            xlHoja1.Cells(liLineas, 12) = RsRRHH(11).Name
            xlHoja1.Cells(liLineas, 13) = RsRRHH(12).Name
            xlHoja1.Cells(liLineas, 14) = RsRRHH(13).Name
            xlHoja1.Cells(liLineas, 15) = RsRRHH(14).Name
    
    '        liLineas = liLineas + 1
    '
    '        xlHoja1.Cells(liLineas, 2) = Mes
    '        xlHoja1.Cells(liLineas, 4) = Mes
    '        xlHoja1.Cells(liLineas, 6) = "Diciembre"
    
            liLineas = liLineas + 2
    
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas - 3, 11)).Font.Bold = True

    '******************************************************
            Do While Not RsRRHH.EOF
                xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 15)).NumberFormat = "###,###,###,##0.00"
                xlHoja1.Cells(liLineas, 1) = RsRRHH(0)
                xlHoja1.Cells(liLineas, 2) = RsRRHH(1)
                xlHoja1.Cells(liLineas, 3) = RsRRHH(2)
                xlHoja1.Cells(liLineas, 4) = RsRRHH(3)
                xlHoja1.Cells(liLineas, 5) = RsRRHH(4)
                xlHoja1.Cells(liLineas, 6) = RsRRHH(5)
                xlHoja1.Cells(liLineas, 7) = RsRRHH(6)
                xlHoja1.Cells(liLineas, 8) = RsRRHH(7)
                xlHoja1.Cells(liLineas, 9) = RsRRHH(8)
                xlHoja1.Cells(liLineas, 10) = RsRRHH(9)
                xlHoja1.Cells(liLineas, 11) = RsRRHH(10)
                xlHoja1.Cells(liLineas, 12) = RsRRHH(11)
                xlHoja1.Cells(liLineas, 13) = RsRRHH(12)
                xlHoja1.Cells(liLineas, 14) = RsRRHH(13)
                xlHoja1.Cells(liLineas, 15) = RsRRHH(14)
                liLineas = liLineas + 1
                RsRRHH.MoveNext
            Loop
        End If
        RsRRHH.Close
        ExcelCuadro xlHoja1, 1, 7, 15, liLineas, True, True
    '******************************************************
'        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9

    'Libera los objetos.
'    Dim oDR As New clsSSKAccess
'    oDR.GrabarBitacora frmBNavegador.Usuario, "BALANCE GENERAL " & sTit2, frmBalanceGeneral.Name, frmBNavegador.tvwNav.SelectedItem.Key & "@" & frmBNavegador.lsvLista.SelectedItem.SubItems(2)

        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        xlAplicacion.Application.Visible = True
        xlAplicacion.Windows(1).Visible = True
End Sub

Public Sub Generar_ReportePlanilla(Periodo As String)

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    Dim liLineaInicio As Integer
    Dim liLineas As Integer
    Dim liLineasTemp As Integer
    Dim licont As Integer
    Dim liNro As Integer

    Dim i As Integer
    Dim lnTotal As Double

    Dim glsArchivo As String

    Dim CadTit As String
    Dim CadImp   As String
    Dim lnDivide As Integer
    Dim RsRRHH As New ADODB.Recordset
    
    Dim sTit1 As String
    Dim sTit As String
    Dim sTit2 As String
    Dim sTit3 As String
    Dim sTit4 As String
    Dim nPos As Integer
    Dim Fila As Integer
    Dim Año As Integer
    Dim mes As Integer
    

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Item(1)
    
    Dim lsNomHoja As String
        

        '**************************************************************************
        Año = Right(Periodo, 4)
        mes = Mid(Periodo, 4, 2)
        xlHoja1.Name = Año & Right("00" & mes, 2)
'        MsgBox Año & Right("00" & mes, 2)

         Set RsRRHH = CargaConsultaPlanillas(Año & Right("00" & mes, 2))
         
'         Set dBalance = Nothing

         xlAplicacion.Range("A1:A1").ColumnWidth = 10
         xlAplicacion.Range("B1:B1").ColumnWidth = 40
         xlAplicacion.Range("C1:C1").ColumnWidth = 10
         xlAplicacion.Range("D1:D1").ColumnWidth = 10

'         CadTit = Trim(prs!cDescrip)

         xlHoja1.Cells(1, 1) = gsNomCmac
         xlHoja1.Cells(1, 3) = "Fecha: " & Date
         xlHoja1.Cells(3, 1) = "REPORTE PLANILLA DE INCENTIVO POR PRODUCTIVIDAD DE " & Format(Periodo, "mmmm") & " DEL " & Year(Periodo)

         xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 3)).Font.Bold = True
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 4)).Merge True
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 4)).HorizontalAlignment = xlCenter
         xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 4)).Font.Bold = True
         xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 4)).Font.Bold = True

         liLineas = 6
        If Not RsRRHH.EOF Then
            liLineas = liLineas + 1
            xlHoja1.Cells(liLineas, 1) = RsRRHH(0).Name
            xlHoja1.Cells(liLineas, 2) = RsRRHH(1).Name
            xlHoja1.Cells(liLineas, 3) = RsRRHH(2).Name
            xlHoja1.Cells(liLineas, 4) = RsRRHH(3).Name
    
    '        liLineas = liLineas + 1
    '
    '        xlHoja1.Cells(liLineas, 2) = Mes
    '        xlHoja1.Cells(liLineas, 4) = Mes
    '        xlHoja1.Cells(liLineas, 6) = "Diciembre"
    
            liLineas = liLineas + 2
    
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas - 3, 11)).Font.Bold = True

    '******************************************************
            Do While Not RsRRHH.EOF
                xlHoja1.Range(xlHoja1.Cells(liLineas, 4), xlHoja1.Cells(liLineas, 4)).NumberFormat = "###,###,###,##0.00"
                xlHoja1.Cells(liLineas, 1) = RsRRHH(0)
                xlHoja1.Cells(liLineas, 2) = RsRRHH(1)
                xlHoja1.Cells(liLineas, 3) = RsRRHH(2)
                xlHoja1.Cells(liLineas, 4) = RsRRHH(3)
                liLineas = liLineas + 1
                RsRRHH.MoveNext
            Loop
        End If
        RsRRHH.Close

        ExcelCuadro xlHoja1, 1, 7, 4, liLineas, True, True
    '******************************************************
'        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9

    'Libera los objetos.
'    Dim oDR As New clsSSKAccess
'    oDR.GrabarBitacora frmBNavegador.Usuario, "BALANCE GENERAL " & sTit2, frmBalanceGeneral.Name, frmBNavegador.tvwNav.SelectedItem.Key & "@" & frmBNavegador.lsvLista.SelectedItem.SubItems(2)

        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        xlAplicacion.Application.Visible = True
        xlAplicacion.Windows(1).Visible = True
End Sub

Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal X1 As Currency, ByVal Y1 As Currency, ByVal X2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).BorderAround xlContinuous, xlThin
If lbLineasVert Then
   If X2 <> X1 Then
     xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End If
End If
If lbLineasHoriz Then
    If Y1 <> Y2 Then
        xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
End If
End Sub

Private Function CargarDatosCuadre(ByVal psFecha As String, ByVal psFechaF As String) As ADODB.Recordset
On Error GoTo CargarDatosCuadreErr
    Dim oCon As DConecta, sSQL As String, Rs As New ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = " Select (Select cpersnombre from persona where cperscod  = a.crhperscod ) cPersNombre , a.cCodCta, Sum(nRHExtraMonto) Monto, aa.Monto from rhextraplanilla a " & _
               " Left join (Select ccodcta, Sum(nMontran) Monto from [128.107.2.3].dbcmact01.dbo.trandiariaconsol where dfectran >= '" & Format(psFecha, "yyyymmdd") & "' and dfectran < '" & Format(psFechaF, "yyyymmdd") & "' " & _
               " and ccodcta in (Select cCodCta from rhextraplanilla where crrhhperiodo like '" & Format(psFecha, "yyyymm") & "%' and cPlanillaTpoCod = 'E01') " & _
               " and ccodope in ('201801','202201') group by ccodcta) aa On a.ccodcta = aa.ccodcta " & _
               " where crrhhperiodo like '" & Format(psFecha, "yyyymm") & "%' " & _
               " group by a.ccodcta , a.crhperscod, aa.Monto order by aa.monto "
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    Set CargarDatosCuadre = Rs
    Exit Function
CargarDatosCuadreErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Sub ImprimeRepCuadre(ByVal psFecha As String, ByVal psFechaF As String)
    On Error GoTo ImprimeErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja  As Boolean, sFecha As String

    Dim liLineas As Integer
    
    Dim i As Integer, fdFechaI As String, fdFechaF As String, fdFecha As String
    
    Dim Rs As New ADODB.Recordset
    Dim lsNomHoja As String, sNombre As String
    
    sNombre = "\spooler\RepCuadre_" & Format(psFecha, "yyyymmdd") & "_" & Format(psFechaF, "yyyymmdd")
    
    If Month(psFecha) <> Month(psFechaF) Then
        MsgBox "La fecha de inicio y la de fin de pertenecer al mismo mes", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set Rs = CargarDatosCuadre(psFecha, psFechaF)
    
    If Rs.BOF Or Rs.EOF Then
        MsgBox "No hay datos para mostrar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
        
        
    Set xlHoja1 = xlLibro.Worksheets.Item(1)
    
    xlHoja1.Name = "Reporte_Cuadre"
            
    '**************************************************************************
    
    liLineas = 2
    xlHoja1.Cells(liLineas, 2) = gsNomCmac
    xlHoja1.Cells(liLineas, 5) = "Fecha " & gdFecSis
    
    liLineas = 3
    xlHoja1.Cells(liLineas, 5) = "Usuario " & gsCodUser
    
    liLineas = 6
    xlHoja1.Cells(liLineas, 2) = "Reporte Cuadre al " & psFecha
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
    
    liLineas = 9
    xlHoja1.Cells(liLineas, 2) = "Nombre"
    xlHoja1.Cells(liLineas, 3) = "Cuenta"
    xlHoja1.Cells(liLineas, 4) = "Monto 1"
    xlHoja1.Cells(liLineas, 5) = "Monto 2"
    
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Cells.Interior.Color = &HC0FFFF
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
    liLineas = liLineas + 1
    
'    LblGenerando.Visible = False
'    Barra.Visible = True
'    Barra.max = Rs.RecordCount
'    Barra.value = 0
        
    Do While Not Rs.EOF
        xlHoja1.Cells(liLineas, 2) = Rs("cPersNombre")
        xlHoja1.Cells(liLineas, 3) = Rs("cCodCta")
        xlHoja1.Range(xlHoja1.Cells(liLineas, 5), xlHoja1.Cells(liLineas, 5)).Cells.NumberFormat = "###,###,##0.00"
        xlHoja1.Cells(liLineas, 4) = Rs(2)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 5), xlHoja1.Cells(liLineas, 5)).Cells.NumberFormat = "###,###,##0.00"
        xlHoja1.Cells(liLineas, 5) = Rs(3)
        'Barra.value = Rs.AbsolutePosition
        Rs.MoveNext
        liLineas = liLineas + 1
    Loop
    
    xlHoja1.Cells(liLineas, 2) = "Total"
    i = 4
    Do While i < 6
        xlHoja1.Range(xlHoja1.Cells(liLineas, i), xlHoja1.Cells(liLineas, i)).Formula = "=SUM(" & Chr(64 + i) & "10:" & Chr(64 + i) & liLineas - 1 & ")"
        i = i + 1
    Loop
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Cells.Interior.Color = &HC0FFFF
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 5)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 4)).Merge True
    xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(3, 6)).Font.Bold = True
    
    ExcelCuadro xlHoja1, 2, 9, 5, liLineas, True, True
    
    Screen.MousePointer = 0
'******************************************************

'    xlHoja1.Cells.Select
    'xlHoja1.Cells.NumberFormat = "###,###,##0.00"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
'    xlHoja1.Cells.EntireRow.AutoFit

    'Libera los objetos.
    Set xlHoja1 = Nothing
        
    xlLibro.SaveAs App.path & sNombre
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    Set xlAplicacion = Nothing
    'Barra.Visible = False
    Exit Sub
ImprimeErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub

Private Function CargarDatosSeguro(ByVal pnAño As Integer) As ADODB.Recordset
On Error GoTo CargarDatosSeguroErr
    Dim oCon As DConecta, sSQL As String, Rs As New ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = " Select cRHCod,  (Select cPersNombre From Persona PE where PE.cPersCod = EM.cPersCod) Nombre, Age.cAgeCod, cRHAreaCodOficial, (Select cAreaDescripcion from Areas A where cAreaCod = cRHAreaCodOficial) AreaDescrip  , Age.cAgeDescripcion, Categoria = Case lefT(ca.cRHCargoCod,3)  when '001'  then 'FUNCIONARIO'    when '002' then 'FUNCIONARIO'  when '003' then 'FUNCIONARIO' Else 'EMPLEADO' end, (Select Case When nAgrupacion in (4,5,6) then 'TAREAS BANCARIAS' ELSE 'TAREAS NO BANCARIAS' END from Areas A where cAreaCod = cRHAreaCodOficial) Tpo, " & _
               " (Select (Select cConsDescripcion from constante where nConsCod = 1043 and nConsValor = nAgrupacion) from Areas A where cAreaCod = cRHAreaCodOficial) " & _
               " , cat.crhcargodescripcion  from RRHH EM " & _
               " Inner Join RHCargos CA On EM.cPersCod = CA.cPersCod " & _
               " Inner join RHContrato CO On EM.cPersCod = CO.cPersCod " & _
               " Inner join Agencias Age On CA.crhagenciaCodOficial = Age.cAgeCod " & _
               " Inner Join RHCargosTabla CAT On CAT.cRHCargoCod = CA.cRHCargoCod " & _
               " Where CA.cPersCod In ( " & _
               " Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & pnAño & "12%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,'" & pnAño & "/12/31') = 0 and datediff(day,dcese,'" & pnAño & "/12/31') <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & pnAño & "12%' And cPlanillaCod = 'E01')) " & _
               " Union All " & _
               " Select cPersCod from rrhh Where datediff(month,dingreso,'" & pnAño & "/12/31') = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & pnAño & "12%' And cPlanillaCod = 'E01') " & _
               " Union All " & _
               " Select cPersCod from rrhh Where datediff(month,dingreso,'" & pnAño & "/12/31') = 0 and crhcod like 'E%' and datediff(month,dcese,'" & pnAño & "/12/31') = 0 and datediff(day,dcese,'" & pnAño & "/12/31') = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & pnAño & "12%' And cPlanillaCod = 'E01') " & _
               " Union All " & _
               " Select cPersCod from rrhh Where datediff(month,dcese,'" & pnAño & "/12/31') = 0 and datediff(day,dcese,'" & pnAño & "/12/31') = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & pnAño & "12%' And cPlanillaCod = 'E01')) " & _
               " And cRHCod like 'E%' And CO.cRHContratoNro in (Select Max(CO1.cRHContratoNro) From RHContratoDet CO1   Where CO1.cPersCod = CO.cPersCod And Convert(varchar(10),dRHContratoFecha,112) <= '" & pnAño & "1231' ) And CA.dRHCargoFecha   = (Select Max(CA1.dRHCargoFecha) From RHCargos CA1 Where CA.cPersCod = CA1.cPersCod And convert(varchar(10),CA1.dRHCargoFecha,112) <= '" & pnAño & "1231') " & _
               " Order by cRHCod "
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    Set CargarDatosSeguro = Rs
    Exit Function
CargarDatosSeguroErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Sub ImprimeRepDatosSeguro(ByVal pnAño As Integer)
    On Error GoTo ImprimeErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja  As Boolean, sFecha As String

    Dim liLineas As Integer
    
    Dim i As Integer, fdFechaI As String, fdFechaF As String, fdFecha As String
    
    Dim Rs As New ADODB.Recordset
    Dim lsNomHoja As String, sNombre As String
    
    sNombre = "\spooler\RepDatosSeguro_" & pnAño
    
    Set Rs = CargarDatosSeguro(pnAño)
    
    If Rs.BOF Or Rs.EOF Then
        MsgBox "No hay datos para mostrar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
        
        
    Set xlHoja1 = xlLibro.Worksheets.Item(1)
    
    xlHoja1.Name = "Reporte_Seguro"
            
    '**************************************************************************
    
    liLineas = 2
    xlHoja1.Cells(liLineas, 2) = gsNomCmac '"Caja Municipal de Ahorros y Creditos Trujillo S. A."
    xlHoja1.Cells(liLineas, 8) = "Fecha " & gdFecSis
    xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).HorizontalAlignment = xlRight
    
    liLineas = 3
    xlHoja1.Cells(liLineas, 8) = "Usuario " & gsCodUser
    xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).HorizontalAlignment = xlRight
    
    liLineas = 6
    xlHoja1.Cells(liLineas, 2) = "Reporte Datos Seguro para el Año " & pnAño
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).Merge True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).HorizontalAlignment = xlCenter
    
    liLineas = 9
    xlHoja1.Cells(liLineas, 2) = "Codigo"
    xlHoja1.Cells(liLineas, 3) = "Nombre"
    xlHoja1.Cells(liLineas, 4) = "Agencia"
    xlHoja1.Cells(liLineas, 5) = "Area"
    xlHoja1.Cells(liLineas, 6) = "Categoria"
    xlHoja1.Cells(liLineas, 7) = "Tipo"
    xlHoja1.Cells(liLineas, 8) = "Descripcion Cargo"
    
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).Cells.Interior.Color = &HC0FFFF
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 8)).HorizontalAlignment = xlCenter
    liLineas = liLineas + 1
          
    Do While Not Rs.EOF
        xlHoja1.Cells(liLineas, 2) = Trim(Rs("cRHCod"))
        xlHoja1.Cells(liLineas, 3) = Trim(Rs("Nombre"))
        xlHoja1.Cells(liLineas, 4) = Trim(Rs("cAgeDescripcion"))
        xlHoja1.Cells(liLineas, 5) = Trim(Rs("AreaDescrip"))
        xlHoja1.Cells(liLineas, 6) = Trim(Rs("Categoria"))
        xlHoja1.Cells(liLineas, 7) = Trim(Rs("Tpo"))
        xlHoja1.Cells(liLineas, 8) = Trim(Rs("crhcargodescripcion"))
        Rs.MoveNext
        liLineas = liLineas + 1
    Loop
    
'    xlHoja1.Cells(liLineas, 2) = "Total"
'    i = 4
'    Do While i < 6
'        xlHoja1.Range(xlHoja1.Cells(liLineas, i), xlHoja1.Cells(liLineas, i)).Formula = "=SUM(" & Chr(64 + i) & "10:" & Chr(64 + i) & liLineas - 1 & ")"
'        i = i + 1
'    Loop
'    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Cells.Interior.Color = &HC0FFFF
'    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 8)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(3, 8)).Font.Bold = True
    
    ExcelCuadro xlHoja1, 2, 9, 8, liLineas - 1, True, True
    
    Screen.MousePointer = 0
'******************************************************

'    xlHoja1.Cells.Select
    'xlHoja1.Cells.NumberFormat = "###,###,##0.00"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
'    xlHoja1.Cells.EntireRow.AutoFit

    'Libera los objetos.
    Set xlHoja1 = Nothing
        
    xlLibro.SaveAs App.path & sNombre
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    Set xlAplicacion = Nothing
    'Barra.Visible = False
    Exit Sub
ImprimeErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub



