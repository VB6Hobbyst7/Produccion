Attribute VB_Name = "gContabilidad"
Option Explicit
Public Const gContLibroMayCta = "760010"

Public Sub ValidaBalanceEXCEL(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer)
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim nDeduccion1 As Currency
Dim sValida    As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
Dim oBala As COMNAuditoria.NBalanceCont

'********************************************
Dim fs As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim I               As Integer
Dim glsArchivo      As String
Dim glsArchivo1      As String
Dim lsNomHoja       As String


nUtilidad = 0
nUtilidadMes = 0
Set oBala = New COMNAuditoria.NBalanceCont

If Month(pdFechaIni) > 1 Then
   nUtilidad = oBala.GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If

n5 = oBala.getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = oBala.getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = oBala.getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = oBala.getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = oBala.getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = oBala.getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = oBala.getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)

'69
nRei = oBala.getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
If gsCodCMAC = "102" Then
   nRei = nRei * -1
End If
nDeduccion = oBala.getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion1 = oBala.getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion * -1
nDeduccion1 = nDeduccion1 * -1

If Not lSoloUtilidad Then
  glsArchivo = "C A L C U L O   D E   L A   U T I L I D A D" & " " & " " & Format(gdFecSis, "ddmmyyyy") & "_" & Format(Time(), "HHMMSS") & ".XLS"
   If pnMoneda = 0 Then
      glsArchivo1 = "C O N S O L I D A D O" & " " & "AL " & pdFechaFin
   End If
End If

    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "CalculoDeLaUtilidad"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 18
            xlAplicacion.Range("B1:B1").ColumnWidth = 40
            xlAplicacion.Range("C1:C1").ColumnWidth = 20
            xlAplicacion.Range("D1:D1").ColumnWidth = 20
          
            xlAplicacion.Range("A1:Z100").Font.Size = 9

            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(1, 2) = glsArchivo
            xlHoja1.Cells(2, 2) = glsArchivo1
            'xlHoja1.Cells(3, 2) = "INFORMACION  AL  " & Format(gdFecSis, "dd/mm/yyyy")

            xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(1, 4)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 4)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(1, 4)).Merge True
            'xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 4)).Merge True
            'xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 4)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(1, 4)).HorizontalAlignment = xlCenter
            'xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 4)).HorizontalAlignment = xlCenter
            'xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 4)).HorizontalAlignment = xlCenter
   
                     
            liLineas = 4
            
            
         
            xlHoja1.Cells(liLineas, 2) = "UTILIDAD ACUMULADA AL " & CDate(pdFechaIni) - 1
            xlHoja1.Cells(liLineas, 3) = PrnVal(nUtilidad, 16, 2)
            
            xlHoja1.Cells(liLineas + 1, 2) = "UTILIDAD DEL MES DE " & Mid(pdFechaFin, 4, 10)
            xlHoja1.Cells(liLineas + 1, 3) = PrnVal(nUtilidadMes - nUtilidad, 16, 2)
            
            xlHoja1.Cells(liLineas + 2, 3) = "-----------------------------------"
            
            xlHoja1.Cells(liLineas + 3, 2) = "UTILIDAD ACUMULADA AL " & CDate(pdFechaFin)
            xlHoja1.Cells(liLineas + 3, 3) = PrnVal(nUtilidadMes, 16, 2)
            
            xlHoja1.Cells(liLineas + 4, 2) = "R.E.I. " & Right(pdFechaFin, 4)
            xlHoja1.Cells(liLineas + 4, 3) = PrnVal(nUtilidadMes, 16, 2)
            
            xlHoja1.Cells(liLineas + 5, 3) = "------------------------------------"
            
            xlHoja1.Cells(liLineas + 6, 2) = "UTILIDAD ANTES DE IMPUESTO"
            xlHoja1.Cells(liLineas + 6, 3) = PrnVal(nUtilidadMes + nRei, 16, 2)
            
            xlHoja1.Cells(liLineas + 7, 2) = "PARTICIPACION LABORAL"
            xlHoja1.Cells(liLineas + 7, 3) = PrnVal(nDeduccion, 16, 2)
            
            xlHoja1.Cells(liLineas + 8, 2) = "IMPUESTO A LA RENTA"
            xlHoja1.Cells(liLineas + 8, 3) = PrnVal(nDeduccion1, 16, 2)
            
            xlHoja1.Cells(liLineas + 9, 3) = "-------------------------------------"
            
            xlHoja1.Cells(liLineas + 10, 2) = "UTILIDAD(PERDIDA) NETA AL" & " " & pdFechaFin
            xlHoja1.Cells(liLineas + 10, 3) = PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
        
            'xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas + 10, 3)).Style = "Comma"
If Not lSoloUtilidad Then

   Dim nActivo As Currency
   Dim nPasivo As Currency
   Dim nPatri  As Currency

   'Eliminamos si Existe la Utilidad Acumulada del Mes
   Dim dBalance33 As COMNAuditoria.DbalanceCont
   Set dBalance33 = New COMNAuditoria.DbalanceCont
   dBalance33.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance33.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance33.EjecutaBatch

   nActivo = oBala.getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = oBala.getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = oBala.getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))

   xlHoja1.Range(xlHoja1.Cells(liLineas + 14, 2), xlHoja1.Cells(liLineas + 14, 3)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(liLineas + 15, 2), xlHoja1.Cells(liLineas + 15, 3)).Font.Bold = True
   
   xlHoja1.Cells(liLineas + 14, 2) = " CONSISTENCIA DE CUADRE DEL BALANCE "
   Select Case pnMoneda
      Case 0: xlHoja1.Cells(liLineas + 15, 2) = " ( CONSOLIDADO ) "
      Case 1: xlHoja1.Cells(liLineas + 15, 2) = " ( MONEDA NACIONAL ) "
      Case 2: xlHoja1.Cells(liLineas + 15, 2) = " ( MONEDA EXTRANJERA ) "
   End Select
   xlHoja1.Cells(liLineas + 17, 2) = "ACTIVO"
   xlHoja1.Cells(liLineas + 17, 3) = PrnVal(nActivo, 16, 2)
   
   xlHoja1.Cells(liLineas + 18, 2) = "PASIVO"
   xlHoja1.Cells(liLineas + 18, 4) = PrnVal(nPasivo, 16, 2)
   
   xlHoja1.Cells(liLineas + 19, 2) = "PATRIMONIO"
   xlHoja1.Cells(liLineas + 19, 4) = PrnVal(nPatri, 16, 2)
   
   xlHoja1.Cells(liLineas + 20, 2) = "UTILIDAD (PERDIDA) NETA"
   xlHoja1.Cells(liLineas + 20, 4) = PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
   
   xlHoja1.Cells(liLineas + 21, 3) = "-------------------------------------"
   xlHoja1.Cells(liLineas + 21, 4) = "-------------------------------------"
   
   xlHoja1.Cells(liLineas + 22, 3) = PrnVal(nActivo, 16, 2)
   xlHoja1.Cells(liLineas + 22, 4) = PrnVal(nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
   
   xlHoja1.Cells(liLineas + 23, 3) = "-------------------------------------"
   xlHoja1.Cells(liLineas + 23, 4) = "-------------------------------------"
       
   xlHoja1.Cells(liLineas + 24, 2) = "DIFERENCIA"
   xlHoja1.Cells(liLineas + 24, 3) = PrnVal(nActivo - (nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1), 16, 2)

End If

        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        gFunContab.ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        Set oBala = Nothing
        Set dBalance33 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call gFunContab.CargaArchivo(glsArchivo, App.path & "\SPOOLER\")

End Sub

'Auditoria: Tesoreria
Public Function ImprimeAsientosContables(sMovs As String, prg As Object, stat As Object, sFechas As String, Optional lPrg As Boolean = True) As String
Dim sSql As String
Dim rs As ADODB.Recordset
Dim rsDoc As ADODB.Recordset
Dim n As Integer
Dim nLin As Integer, P As Integer
Dim nTot As Currency
Dim nTotH As Currency
Dim sDoc As String
Dim sTexto As String, sAsiento As String
Dim sMovNro As String, sMovItem As String
Dim nLi As Integer
Dim lsFecha As String
nLin = gnLinPage
Dim oCon As New DConecta
Dim lbSaltaPagina As Boolean
oCon.AbreConexion
sSql = "SELECT a.cMovNro, a.nMovNro, b.nMovItem, a.cMovDesc, b.cCtaContCod, dbo.GetCtaContDesc(b.cCtaContCod,2,1) cCtaContDesc, " _
     & "       b.nMovImporte, ISNULL(me.nMovMEImporte,0) nMovMEImporte, f.nMovTpoCambio nMovTpoCambio, a.nMovFlag, a.nMovEstado " _
     & "FROM   Mov a LEFT JOIN MovTpoCambio f ON f.nMovNro = a.nMovNro " _
     & "             JOIN MovCta   b  ON b.nMovNro = a.nmovNro " _
     & "        LEFT JOIN MovME    me ON me.nMovNro = b.nMovNro and me.nMovItem = b.nMovItem " _
     & "WHERE  a.cMovNro IN (" & sMovs & ") " _
     & "ORDER BY LEFT(a.cMovNro,8), RIGHT(a.cMovNro,4), a.cMovNro, b.nMovItem "
Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
   MsgBox "No se seleccionaron Movimientos a Imprimir", vbInformation, "Aviso"
   Exit Function
End If
If lPrg Then
   prg.Min = 0
   prg.Max = rs.RecordCount
End If
CON = PrnSet("C+")
BON = PrnSet("B+")
COFF = PrnSet("C-")
BOFF = PrnSet("B-")

sTexto = ""
Do While Not rs.EOF
   If lPrg Then
      prg.value = rs.Bookmark
      stat.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   End If
   lbSaltaPagina = False
   If lsFecha <> Left(rs!cMovNro, 8) Then
      lsFecha = Left(rs!cMovNro, 8)
      lbSaltaPagina = True
   End If
   sMovNro = rs!cMovNro
   gsGlosa = Replace(rs!cMovDesc, Chr(13) & oImpresora.gPrnSaltoLinea, " ")
   nLi = 0
   sAsiento = CON & " Nro.Mov.: " & sMovNro
   If Not IsNull(rs!nMovTpoCambio) Then
      sAsiento = sAsiento & Space(10) & "  T.Cambio Mercado: " & Format(rs!nMovTpoCambio, "##,###,#00.000") & "      T.Cambio Fijo : " & Format(gnTipCambio, "##,###,#00.000") & oImpresora.gPrnSaltoLinea
   Else
      sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
   End If
   nLi = nLi + 1
   sSql = "SELECT d.cDocAbrev, md.cDocNro, md.dDocFecha  " _
        & "FROM MovDoc md LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
        & "WHERE md.nMovNro = " & rs!nmovnro _
        & " union " _
        & " SELECT  d.cDocAbrev, md.cDocNro, md.dDocFecha   " _
        & " FROM    movref  mr " _
        & "         join movdoc md on md.nmovnro = mr.nMovNroRef " _
        & "         left join documento d on d.nDocTpo = md.nDocTpo " _
        & "         JOIN MOV M ON M.NMOVNRO = MR.nMovNroRef " _
        & " WHERE   mr.nMovNro = " & rs!nmovnro & " AND M.NMOVFLAG=0 "
        
   Set rsDoc = oCon.CargaRecordSet(sSql)
   If Not rsDoc.EOF Then
      sDoc = " Documentos : "
      Do While Not rsDoc.EOF
         sDoc = sDoc & ImpreFormat(rsDoc!cDocAbrev & "-" & rsDoc!cDocNro, 20) & rsDoc!dDocFecha & Space(5)
         rsDoc.MoveNext
      Loop
      sAsiento = sAsiento & sDoc & oImpresora.gPrnSaltoLinea
      nLi = nLi + 1
   End If
   sAsiento = sAsiento & ImpreGlosaTeso(gsGlosa, gnColPage * 1.5, " Glosa : ", , , nLi) & COFF
   nTot = 0
   nTotH = 0
   Do While rs!cMovNro = sMovNro
      sMovItem = rs!nMovItem
      sAsiento = sAsiento & CON & Justifica(rs!nMovItem, 4) & " " & Mid(rs!cCtaContCod & Space(22), 1, 22) & " " & Mid(rs!cCtaContDesc & Space(46), 1, 46) _
          & Right(Space(14) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, gsFormatoNumeroView), ""), 14) & " " _
          & Right(Space(14) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, gsFormatoNumeroView), ""), 14) _
          & Right(Space(14) & IIf(rs!nMovMEImporte > 0, Format(rs!nMovMEImporte, gsFormatoNumeroView), ""), 14) & " " _
          & Right(Space(14) & IIf(rs!nMovMEImporte < 0, Format(rs!nMovMEImporte * -1, gsFormatoNumeroView), ""), 14) _
          & COFF & oImpresora.gPrnSaltoLinea
      If rs!nMovImporte > 0 Then
         nTot = nTot + Val(rs!nMovImporte)
      Else
         nTotH = nTotH + Val(rs!nMovImporte) * -1
      End If
      nLi = nLi + 1
      If lPrg Then
         prg.value = rs.Bookmark
         stat.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
      End If
      rs.MoveNext
      If rs.EOF Then
         Exit Do
      End If
   Loop
   LineaT sAsiento, CON & String(72, "-") & Right(String(14, "-") & Format(nTot, gsFormatoNumeroView), 14) & "-" & Right(String(14, "-") & Format(nTotH, gsFormatoNumeroView), 14) & BOFF & COFF, , nLi
   If nLi + nLin + 3 > gnLinPage Or lbSaltaPagina Then
      If P > 0 Then sTexto = sTexto & oImpresora.gPrnSaltoPagina
      P = P + 1
      LineaT sTexto, Justifica(gsNomCmac, 55) & gdFecSis & " - " & Format(Time, "hh:mm:ss")
      LineaT sTexto, Space(72) & "Pag. " & Format(P, "000")
      LineaT sTexto, BON & Centra(" A S I E N T O S   C O N T A B L E S ", gnColPage)
      LineaT sTexto, Centra(" M O N E D A   " & IIf(gsSimbolo = gcME, "E X T R A N J E R A ", "N A C I O N A L "), gnColPage)
      LineaT sTexto, Centra(sFechas, gnColPage) & BOFF & CON
      LineaT sTexto, "========================================================================================================" & IIf(gsSimbolo = gcME, "==========================", "")
      LineaT sTexto, "Item C U E N T A     C O N T A B L E                                             DEBE          HABER    " & IIf(gsSimbolo = gcME, "       DEBE         HABER ", "")
      LineaT sTexto, "     Código        Descripción                                                   M.N.          M.N.     " & IIf(gsSimbolo = gcME, "       M.E.         M.E.  ", "")
      LineaT sTexto, "--------------------------------------------------------------------------------------------------------" & IIf(gsSimbolo = gcME, "--------------------------", "") & COFF
      nLin = 9
   End If
   nLin = nLin + nLi
   sTexto = sTexto & sAsiento
Loop
   LineaT sTexto, CON & "========================================================================================================" & IIf(gsSimbolo = gcME, "==========================", "")
ImprimeAsientosContables = PrnSet("MI", 4) & sTexto
End Function

'Public Function ImpreGlosa(psGlosa As String, pnColPage As Integer, Optional psTitGlosa As String = "  GLOSA      : ", Optional pnCols As Integer = 0, Optional lbEnterFinal As Boolean = True, Optional ByRef nLin As Integer = 0) As String
'Dim sImpre As String
'Dim sTexto As String, N As Integer
'Dim nLen As Integer
'  nLen = Len(psTitGlosa)
'  sTexto = JustificaTexto(psGlosa, IIf(pnCols = 0, pnColPage, pnCols) - nLen)
'  sImpre = psTitGlosa
'  N = 0
'  Do While True
'     N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
'     If N > 0 Then
'        sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
'        sTexto = Mid(sTexto, N + 1, Len(sTexto))
'        nLin = nLin + 1
'     End If
'     If N = 0 Then
'        sImpre = sImpre & Justifica(sTexto, IIf(pnCols = 0, pnColPage, pnCols) - nLen) & IIf(lbEnterFinal, oImpresora.gPrnSaltoLinea, "")
'        If lbEnterFinal Then
'            nLin = nLin + 1
'        End If
'        Exit Do
'     End If
'  Loop
'  ImpreGlosa = sImpre
'End Function

Public Function LineaT(psVarImpre As String, psTexto As String, Optional pnLineas As Integer = 1, Optional ByRef pnLinCnt As Integer = 0) As String
Dim K As Integer
psVarImpre = psVarImpre & psTexto
For K = 1 To pnLineas
   psVarImpre = psVarImpre & oImpresora.gPrnSaltoLinea
   pnLinCnt = pnLinCnt + 1
Next
End Function
