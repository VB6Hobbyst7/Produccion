Attribute VB_Name = "gContabilidad"
Option Explicit
Dim sSql As String
Dim rs   As ADODB.Recordset

Public Enum gRiesgosTpoCalculo
   gRiesgosTpoTasa = 1
   gRiesgosTpoVencimiento = 2
   gRiesgosEncaje_BCR = 3
   gRiesgosPlaza_Cheque = 4
   gRiesgosSeries = 5
   gRiesgosFecha = 6
   gRiesgosFormulaAcumula = 7
   gRiesgosFormula = 8
   gRiesgosTotales = 9
   gRiesgosPorcentualSegunCriterio = 10
   gRiesgosEstadisctico = 11
End Enum

'Para Pasar a Clases

'*************************************************************************
'*          VARIABLES PARA EL REPORTE DE TESORERIA                       *
'*************************************************************************
'RATIOS DE LIQUIDEZ **********************
    '***************ACTIVOS LIQUIDOS ************************
    Global lsActCaja(2) As String
    Global lsActBancos(2) As String
    '***************PASIVOS DE CORTO PLAZO ****************************
    Global lsPasOblig(2) As String
    Global lsPasDepAho(2) As String
    Global lsPasDepPlazo(2) As String
    Global lsPasAdeudados(2) As String
'ENCAJE **************************************
    ' ********** TOTAL OBLIGACONES SUJETAS A ENCAJE CONSOLIDADO A NIVEL NACIONAL
    Global lsEncPlazo(2) As String
    Global lsEncAhorros(2) As String
    '*********** POSICION DE ENCAJE
    Global lsPosEncExig(2) As String
    Global lsPosFondosCaja(2) As String
    Global lsPosEncFondosCaja(2) As String
    Global lsPosEncFondosBCRP(2) As String
    Global lsPosEnResultados(2) As String
    Global lsPosEncAcumulada(2) As String
    '*********** CHEQUES DEDUCIDO DEL TOTAL DE OBLIGACIONES
    Global lsChqPlazoHasta30dias(2) As String
    Global lsChqPlazo(2) As String
    Global lsChqAhorros(2) As String
    
    Global lsDepBcoNacion As String
    
    Global gnCuotasNro As Long
    Global gnTasaInteres As Currency
    Global gnDiasVenc As Long

Public Function GetSaldoCtaClase(cCta As String, dFecha As Date, pnMoneda As Integer) As Currency
Dim sSql As String, rs As ADODB.Recordset
Dim oCon As DConecta
     
sSql = "SELECT ISNULL(SUM(a.nCtaSaldoImporte),0) as nSaldo, ISNULL(SUM(a.nCtaSaldoImporteME),0) as nSaldoME " _
     & "FROM  CtaSaldo a " _
     & "WHERE cCtaContcod LIKE '" & cCta & "%' " _
     & "AND a.dCtaSaldoFecha = ( SELECT MAX(b.dCtaSaldoFecha) " _
     & "                          FROM  CtaSaldo b " _
     & "                          WHERE b.cCtaContCod = a.cCtaContCod and b.dCtaSaldoFecha <= '" & Format(dFecha, gsFormatoFecha) & "')"
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If Not rs.EOF Then
   If pnMoneda = 1 Then
      GetSaldoCtaClase = rs!nSaldo
   Else
      GetSaldoCtaClase = rs!nSaldoME
   End If
End If

oCon.CierraConexion
Set oCon = Nothing
End Function

'****************************************************

Public Function ObjetoExiste(psObjetoCod As String) As String
Dim oObj As DObjeto
Set oObj = New DObjeto
Set rs = oObj.CargaObjeto(psObjetoCod)
Set oObj = Nothing
If rs.EOF Then
   MsgBox "Objeto no existe...Por favor verificar", vbInformation, "¡Advertencia!"
   RSClose rs
   Exit Function
End If
ObjetoExiste = rs!cObjetoDesc
RSClose rs
End Function

Public Sub LlenaComboConstante(psConsCod As ConstanteCabecera, psCombo As ComboBox)
Dim lsRs As New ADODB.Recordset
Dim clsCons As New DConstantes
Set lsRs = clsCons.CargaConstante(psConsCod)
If psConsCod = gViaticosDestino Then 'PASI 20140401 TI-ERS050-2014
    Dim oConstSist As New NConstSistemas
    Set lsRs = oConstSist.LeeRutasViaticos()
    If Not RSVacio(lsRs) Then
        psCombo.Clear
        Do While Not lsRs.EOF
            psCombo.AddItem Trim(lsRs(0)) & space(100) & Trim(lsRs(1))
            lsRs.MoveNext
        Loop
    End If
Else
    Set lsRs = clsCons.CargaConstante(psConsCod)
    If Not RSVacio(lsRs) Then
        psCombo.Clear
        Do While Not lsRs.EOF
            psCombo.AddItem Trim(lsRs(2)) & space(100) & Trim(lsRs(1))
            lsRs.MoveNext
        Loop
    End If
End If
RSClose lsRs
Set clsCons = Nothing
End Sub

Public Function BuscaCombo(sValor As String, Cbo As ComboBox) As Integer
Dim N As Integer
BuscaCombo = -1
For N = 0 To Cbo.ListCount - 1
    If Trim(Right(Cbo.List(N), Len(sValor))) = Trim(sValor) Then 'JIPR AGREGÒ TRIM() EN AMBOS CASOS 20180816
      BuscaCombo = N
      Exit For
   End If
Next
End Function

Public Function LeeTpoCambio(psFecha As String, Optional TpoCambio As TipoCambio = TCFijoDia) As Currency
Dim clsTC As New nTipoCambio
On Error GoTo LeeTpoCambioErr
   LeeTpoCambio = clsTC.EmiteTipoCambio(psFecha, TCFijoDia)
   Set clsTC = Nothing
Exit Function
LeeTpoCambioErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function


Public Function ValidaOrdenPagoCuenta(ByVal psCuentaAho As String, ByVal psNroDoc As String, ByVal pnMonto As Currency) As Boolean
Dim clsMant As NCapMantenimiento
Dim nEstadoOP As Long
Dim rsOP As ADODB.Recordset
Set rsOP = New ADODB.Recordset
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
ValidaOrdenPagoCuenta = True
If clsCap.EsOrdenPagoEmitida(psCuentaAho, CLng(psNroDoc)) Then
    Set clsMant = New NCapMantenimiento
    Set rsOP = clsMant.GetDatosOrdenPago(psCuentaAho, CLng(psNroDoc))
    Set clsMant = Nothing
    If Not (rsOP.EOF And rsOP.BOF) Then
        nEstadoOP = rsOP("cEstado")
        If nEstadoOP = gCapOPEstAnulada Or nEstadoOP = gCapOPEstCobrada Or nEstadoOP = gCapOPEstExtraviada Or nEstadoOP = gCapOPEstRechazada Then
            MsgBox "Orden de Pago N° " & psNroDoc & " " & rsOP("cDescripcion"), vbInformation, "Aviso"
            rsOP.Close
            Set rsOP = Nothing
            ValidaOrdenPagoCuenta = False
            Exit Function
        ElseIf rsOP("cEstado") = gCapOPEstCertifiCada Then
            If pnMonto <> rsOP("nMonto") Then
                MsgBox "Orden de Pago Certificada. Monto No Coincide con Monto de Certificación", vbInformation, "Aviso"
                rsOP.Close
                Set rsOP = Nothing
                Exit Function
                ValidaOrdenPagoCuenta = False
            End If
        End If
    End If
    rsOP.Close
    Set rsOP = Nothing
Else
    MsgBox "Orden de Pago No ha sido emitida para esta cuenta", vbInformation, "Aviso"
    ValidaOrdenPagoCuenta = False
    Exit Function
End If
End Function

Public Sub ProgressShow(oPrg As Object, oForm As Object, Optional psCaption As eCaptionStyle = eCap_CaptionOnly, Optional pbCallOut As Boolean = False)

    oPrg.CaptionSyle = psCaption
    oPrg.Max = 1
    oPrg.Color1 = vbWhite
    oPrg.Color2 = vbBlue
    oPrg.ShowForm oForm
    If Not pbCallOut Then
        oPrg.Visible = True
        oPrg.Top = oPrg.GetTop(oForm)
        oPrg.Left = oPrg.GetLeft(oForm)
    End If
    oPrg.Progress 0
    DoEvents
End Sub

Public Sub ProgressClose(oPrg As Object, oForm As Object, Optional pbCallOut As Boolean = False)
If Not pbCallOut Then
    oPrg.Visible = False
End If
    oPrg.CloseForm oForm
End Sub

Public Function ValidaFechaContab(pdFecha As String, pdFecsis As Date, Optional pbVerCierreContab As Boolean = True) As Boolean
ValidaFechaContab = False
   If ValidaFecha(Format(pdFecha, gsFormatoFechaView)) <> "" Then
      MsgBox " Fecha no Válida... ", vbInformation, "Aviso"
      Exit Function
   Else
      If CDate(Trim(pdFecha)) > pdFecsis Then
         MsgBox "Fecha no puede ser mayor a fecha Actual...", vbInformation, "¡Aviso!"
         Exit Function
      End If
      If Year(pdFecsis) - Year(pdFecha) > 1 Then
         MsgBox "Fecha de Años anteriores no permitida...", vbInformation, "¡Aviso!"
         Exit Function
      End If
   End If
   If pbVerCierreContab Then
    Dim oCont As New NContFunciones
    If Not oCont.PermiteModificarAsiento(Format(pdFecha, gsFormatoMovFecha), False) Then
       Set oCont = Nothing
       MsgBox "Mes Contable ya cerrado. Fecha de Operación no Permitida", vbInformation, "!Aviso!"
       Exit Function
    End If
    Set oCont = Nothing
   End If
   
ValidaFechaContab = True
End Function

Public Function ImprimeAsientosContables(sMovs As String, prg As Object, stat As Object, sFechas As String, Optional lPrg As Boolean = True) As String
Dim sSql As String
Dim rs As ADODB.Recordset
Dim rsDoc As ADODB.Recordset
Dim N As Integer
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
      sAsiento = sAsiento & space(10) & "  T.Cambio Mercado: " & Format(rs!nMovTpoCambio, "##,###,#00.000") & "      T.Cambio Fijo : " & Format(gnTipCambio, "##,###,#00.000") & oImpresora.gPrnSaltoLinea
   Else
      sAsiento = sAsiento & oImpresora.gPrnSaltoLinea
   End If
   nLi = nLi + 1
   sSql = "SELECT d.cDocAbrev, md.cDocNro, md.dDocFecha  " _
        & "FROM MovDoc md LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
        & "WHERE md.nMovNro = " & rs!nMovNro _
        & " union " _
        & " SELECT  d.cDocAbrev, md.cDocNro, md.dDocFecha   " _
        & " FROM    movref  mr " _
        & "         join movdoc md on md.nmovnro = mr.nMovNroRef " _
        & "         left join documento d on d.nDocTpo = md.nDocTpo " _
        & "         JOIN MOV M ON M.NMOVNRO = MR.nMovNroRef " _
        & " WHERE   mr.nMovNro = " & rs!nMovNro & " AND M.NMOVFLAG=0 "
        
   Set rsDoc = oCon.CargaRecordSet(sSql)
   If Not rsDoc.EOF Then
      sDoc = " Documentos : "
      Do While Not rsDoc.EOF
         sDoc = sDoc & ImpreFormat(rsDoc!cDocAbrev & "-" & rsDoc!cDocNro, 20) & rsDoc!dDocFecha & space(5)
         rsDoc.MoveNext
      Loop
      sAsiento = sAsiento & sDoc & oImpresora.gPrnSaltoLinea
      nLi = nLi + 1
   End If
   sAsiento = sAsiento & ImpreGlosa(gsGlosa, gnColPage * 1.5, " Glosa : ", , , nLi) & COFF
   nTot = 0
   nTotH = 0
   Do While rs!cMovNro = sMovNro
      sMovItem = rs!nMovItem
      sAsiento = sAsiento & CON & Justifica(rs!nMovItem, 4) & " " & Mid(rs!cCtaContCod & space(22), 1, 22) & " " & Mid(rs!cCtaContDesc & space(46), 1, 46) _
          & Right(space(14) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, gsFormatoNumeroView), ""), 14) & " " _
          & Right(space(14) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, gsFormatoNumeroView), ""), 14) _
          & Right(space(14) & IIf(rs!nMovMEImporte > 0, Format(rs!nMovMEImporte, gsFormatoNumeroView), ""), 14) & " " _
          & Right(space(14) & IIf(rs!nMovMEImporte < 0, Format(rs!nMovMEImporte * -1, gsFormatoNumeroView), ""), 14) _
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
   Linea sAsiento, CON & String(72, "-") & Right(String(14, "-") & Format(nTot, gsFormatoNumeroView), 14) & "-" & Right(String(14, "-") & Format(nTotH, gsFormatoNumeroView), 14) & BOFF & COFF, , nLi
   If nLi + nLin + 3 > gnLinPage Or lbSaltaPagina Then
      If P > 0 Then sTexto = sTexto & oImpresora.gPrnSaltoPagina
      P = P + 1
      Linea sTexto, Justifica(gsNomCmac, 55) & gdFecSis & " - " & Format(Time, "hh:mm:ss")
      Linea sTexto, space(72) & "Pag. " & Format(P, "000")
      Linea sTexto, BON & Centra(" A S I E N T O S   C O N T A B L E S ", gnColPage)
      Linea sTexto, Centra(" M O N E D A   " & IIf(gsSimbolo = gcME, "E X T R A N J E R A ", "N A C I O N A L "), gnColPage)
      Linea sTexto, Centra(sFechas, gnColPage) & BOFF & CON
      Linea sTexto, "========================================================================================================" & IIf(gsSimbolo = gcME, "==========================", "")
      Linea sTexto, "Item C U E N T A     C O N T A B L E                                             DEBE          HABER    " & IIf(gsSimbolo = gcME, "       DEBE         HABER ", "")
      Linea sTexto, "     Código        Descripción                                                   M.N.          M.N.     " & IIf(gsSimbolo = gcME, "       M.E.         M.E.  ", "")
      Linea sTexto, "--------------------------------------------------------------------------------------------------------" & IIf(gsSimbolo = gcME, "--------------------------", "") & COFF
      nLin = 9
   End If
   nLin = nLin + nLi
   sTexto = sTexto & sAsiento
Loop
   Linea sTexto, CON & "========================================================================================================" & IIf(gsSimbolo = gcME, "==========================", "")
ImprimeAsientosContables = PrnSet("MI", 4) & sTexto
End Function

Public Function ValidaMigracion(ppdfecha As Date) As String
Dim lsTexto As String
Dim NMigra As New DGeneral
Dim oAge As New DActualizaDatosArea
Dim rAge As New ADODB.Recordset
Dim I As Integer
     
'esDomingoOFeriado
  
     
    lsTexto = ""
    Set rAge = oAge.GetAgencias(, False, True)
    Do While Not rAge.EOF
        For I = 1 To 2
            'if esdomingooferiado(ppdfecha) = True And (rAge!Codigo <> "01" And rAge!Codigo <> "07") Then
            
            'Else
                If NMigra.GetMigroAsientoAgencia(ppdfecha, I, rAge!Codigo, gdFecSis) = False Then
                    lsTexto = lsTexto & space(5) & "** [" & ppdfecha & "] Agencia " & rAge!Codigo & " - " & rAge!Descripcion
                    If I = 1 Then
                        lsTexto = lsTexto & space(1) & "[MON NAC]" & Chr(10)
                    ElseIf I = 2 Then
                        lsTexto = lsTexto & space(1) & "[MON EXT]" & Chr(10)
                    End If
                End If
            'End If
        Next
        rAge.MoveNext
    Loop
    rAge.Close
    Set rAge = Nothing
     
ValidaMigracion = lsTexto
 
End Function





Public Sub ValidaBalanceEXCEL(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer)
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim nDeduccion1 As Currency
Dim sValida    As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
Dim oBal As NBalanceCont

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
Set oBal = New NBalanceCont

If Month(pdFechaIni) > 1 Then
   nUtilidad = oBal.GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If

n5 = oBal.getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = oBal.getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = oBal.getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = oBal.getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = oBal.getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = oBal.getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = oBal.getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)

'69
nRei = oBal.getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
If gsCodCMAC = "102" Then
   nRei = nRei * -1
End If
nDeduccion = oBal.getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion1 = oBal.getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
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
   Dim dBalance As New DbalanceCont
   dBalance.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance.EjecutaBatch

   nActivo = oBal.getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = oBal.getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = oBal.getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))

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

   'xlHoja1.Range(xlHoja1.Cells(liLineas + 17, 3), xlHoja1.Cells(liLineas + 24, 4)).Style = "Comma"
End If
'ValidaBalance = sValida

       ' ExcelCuadro xlHoja1, 2, 6, 3, liLineas - 1
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        'Cierra el libro de trabajo
        'xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
        'xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")

End Sub

Public Sub ImprimeBalanceFormaABExcel(xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, xlAplicacion As Excel.Application, psOpeCod As String, pdFecha As Date, psBalanceCate As String, pbSoles As String, psEmpresa As String, psAgenciaCod As String)


Dim lnDivide As Integer
Dim prs As ADODB.Recordset
   
'Dim cMES As String
Dim cnomhoja As String
Dim liLineas As Long
Dim nReg As Integer
Dim lnTipoCambio As Currency
Dim glsArchivo As String
    
If pbSoles Then
    lnDivide = 1
Else
    lnDivide = 1000
End If
    
    
   
'**********************************************

                '************* Hoja ****************
    'cnomhoja = "Balance General"
    cnomhoja = "ESTADO DE SITUACIÓN FINANCIERA"
    
    Call ExcelAddHoja(cnomhoja, xlLibro, xlHoja1)
    
    
           
    xlAplicacion.Range("A1:A1").ColumnWidth = 40
    xlAplicacion.Range("B1:B1").ColumnWidth = 20
    xlAplicacion.Range("C1:C1").ColumnWidth = 20
    xlAplicacion.Range("D1:D1").ColumnWidth = 20
    xlAplicacion.Range("E1:E1").ColumnWidth = 20
    xlAplicacion.Range("A1:Z1").Font.Size = 8
    

 Dim dBalance As New NBalanceCont
    Set prs = dBalance.CargaBalanceGeneral(psOpeCod, Format(pdFecha, "yyyymmdd"), psBalanceCate)
    Set dBalance = Nothing
     
    If Not (prs.EOF And prs.BOF) Then
        liLineas = 1
        
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).Font.Bold = True
        
        xlHoja1.Cells(liLineas, 3) = " ESTADO DE SITUACIÓN FINANCIERA" & IIf(psAgenciaCod = "", "", " AGENCIA - " & psAgenciaCod)
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).Merge True
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).HorizontalAlignment = xlCenter
            
        liLineas = liLineas + 1
        'xlHoja1.Cells(liLineas, 3) = psEmpresa
        xlHoja1.Cells(liLineas, 3) = "CMAC MAYNAS S.A."
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).Merge True
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).HorizontalAlignment = xlCenter
        
        liLineas = liLineas + 1
        xlHoja1.Cells(liLineas, 3) = "ESTADO DE SITUACIÓN FINANCIERA :" & pdFecha
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).Merge True
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).HorizontalAlignment = xlCenter
        
        liLineas = liLineas + 1
        If pbSoles Then
        xlHoja1.Cells(liLineas, 4) = "(Expresado en  Nuevos Soles)"
        Else
        xlHoja1.Cells(liLineas, 4) = "(Expresado en Miles de Nuevos Soles)"
        End If
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).Merge True
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 4)).HorizontalAlignment = xlCenter
        
        liLineas = liLineas + 3
        xlHoja1.Cells(liLineas, 1) = "ACTIVO"
        xlHoja1.Cells(liLineas, 2) = "Moneda Nacional"
        xlHoja1.Cells(liLineas, 3) = "Equivalente en M.E."
        xlHoja1.Cells(liLineas, 4) = "Total"
        'xlHoja1.Cells(liLineas, 5) = "Total Ajustado por Inflacion"
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
        liLineas = liLineas + 1
          
        
        Do While Not prs.EOF
                                
                xlHoja1.Cells(liLineas, 1) = RTrim(prs!cDescrip)
                xlHoja1.Cells(liLineas, 2) = Format(Round(prs!nMN / lnDivide, 2), gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 3) = Format(Round(prs!nME / lnDivide, 2), gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 4) = Format(Round(prs!nTotal / lnDivide, 2), gsFormatoNumeroView)
                'xlHoja1.Cells(liLineas, 5) = Format(Round(prs!nTotAj / lnDivide, 2), gsFormatoNumeroView)
                liLineas = liLineas + 1
            
            If Trim(prs!cDescrip) = "TOTAL ACTIVO" Then
                liLineas = liLineas + 2
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
                
                xlHoja1.Cells(liLineas, 1) = "PASIVO"
                xlHoja1.Cells(liLineas, 2) = "Moneda Nacional"
                xlHoja1.Cells(liLineas, 3) = "Equivalente en M.E."
                xlHoja1.Cells(liLineas, 4) = "Total"
                'xlHoja1.Cells(liLineas, 5) = "Total Ajustado por Inflacion"
                liLineas = liLineas + 1
            ElseIf Trim(prs!cDescrip) = "TOTAL PASIVO" Then
                liLineas = liLineas + 2
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 5)).Font.Bold = True
            
                xlHoja1.Cells(liLineas, 1) = "GANANCIAS Y PERDIDAS"
                xlHoja1.Cells(liLineas, 2) = "Moneda Nacional"
                xlHoja1.Cells(liLineas, 3) = "Equivalente en M.E."
                xlHoja1.Cells(liLineas, 4) = "Total"
                'xlHoja1.Cells(liLineas, 5) = "Total Ajustado por Inflacion"
                liLineas = liLineas + 1
                
            End If
            prs.MoveNext
        Loop
        
   End If
   prs.Close
End Sub

Public Function ImprimeBalanceSectorial(psFecha As String, psEmpresa As String, lbArchivo As Boolean)
Dim CadImp As String
Dim Cont As Integer
Dim P    As Integer
Dim lsFormatoNumero As String

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
Dim prs As ADODB.Recordset


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
            lsNomHoja = "Balance Sectorial"
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

            xlAplicacion.Range("A1:A1").ColumnWidth = 30
            xlAplicacion.Range("B1:B1").ColumnWidth = 55
            xlAplicacion.Range("C1:C1").ColumnWidth = 20
            xlAplicacion.Range("D1:D1").ColumnWidth = 20
            xlAplicacion.Range("E1:E1").ColumnWidth = 20
          
            xlAplicacion.Range("A1:Z100").Font.Size = 9

            xlHoja1.Cells(1, 1) = psEmpresa
            xlHoja1.Cells(2, 2) = "BALANCE SECTORIAL POR AGENTES ECONÓMICOS"
            glsArchivo = xlHoja1.Cells(2, 2) & " " & Format(gdFecSis, "ddmmyyyy") & "_" & Format(Time(), "HHMMSS") & ".XLS"
            xlHoja1.Cells(3, 2) = "AL" & " " & Format(CDate(psFecha), gsFormatoFechaView)
            xlHoja1.Cells(4, 2) = "(Saldos Expresados en Miles de Nuevos Soles)" 'MARG ERS044-2016
            xlHoja1.Cells(4, 2) = "(Saldos Expresados en " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
            
            xlHoja1.Cells(6, 1) = "Codigo de Cuentas"
            xlHoja1.Cells(6, 2) = "Descripcion de la Cuenta"
            xlHoja1.Cells(6, 3) = "Moneda Nacional"
            xlHoja1.Cells(6, 5) = "Moneda"
            
            xlHoja1.Cells(7, 3) = "Ajustado"
            xlHoja1.Cells(7, 4) = "Historico"
            xlHoja1.Cells(7, 5) = "Extranjera"
            
            
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 5)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 5)).Merge True
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 5)).Merge True
            xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 5)).Merge True
            xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 4)).Merge True
            
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 4)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(7, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(7, 4), xlHoja1.Cells(7, 4)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(7, 5), xlHoja1.Cells(7, 5)).HorizontalAlignment = xlCenter
   
    
    Dim dBalance As New NBalanceCont
    
    Set prs = dBalance.CargaBalanceSectorial(psFecha)
    'Cont = 60
    'P = 0
    liLineas = 8
      Do While Not prs.EOF
        If lbArchivo Then
                xlHoja1.Cells(liLineas, 1) = "'" & CStr(Trim(prs!cCodigo))
                xlHoja1.Cells(liLineas, 2) = Format(prs!nMNAj, lsFormatoNumero)
                xlHoja1.Cells(liLineas, 3) = Format(prs!nMNHist, lsFormatoNumero)
                xlHoja1.Cells(liLineas, 4) = Format(prs!nME, lsFormatoNumero)
                'xlHoja1.Cells(liLineas, 5) = Format(Round(prs!nTotAj / lnDivide, 2), gsFormatoNumeroView)
                liLineas = liLineas + 1
                     
         Else
                xlHoja1.Cells(liLineas, 1) = "'" & CStr(Trim(prs!cCodigo))
                xlHoja1.Cells(liLineas, 2) = Mid(prs!cDescrip + space(60), 1, 60)
                xlHoja1.Cells(liLineas, 3) = Format(prs!nMNAj, lsFormatoNumero)
                xlHoja1.Cells(liLineas, 4) = Format(prs!nMNHist, lsFormatoNumero)
                xlHoja1.Cells(liLineas, 5) = Format(prs!nME, lsFormatoNumero)
                liLineas = liLineas + 1
                   
         End If
         prs.MoveNext
      Loop
    prs.Close: Set prs = Nothing
    
    ' ExcelCuadro xlHoja1, 2, 6, 3, liLineas - 1
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        'Cierra el libro de trabajo
        'xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
        'xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")

    'ImprimeBalanceSectorial = CadImp
    'CadImp = ""
End Function

Public Function DevolverSaldoTC(psCtaContCod As String, pnTipoCambioFijo As Currency, pnTipoCambioVenta As Currency, pnTipoCambioCompra As Currency, pnSaldo As Currency, psMoneda As String) As Currency
    DevolverSaldoTC = pnSaldo
End Function

