VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCajaGenImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String
Dim lsEmprLogo As String
Type tBilletajes
    lsBilletes As Currency
    lnMontoBil As Currency
    lsMonedas As Currency
    lnMontoMon As Currency
End Type

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()

'VARIABLES PARA ENCAJE
Dim laSdoEnc() As String, laTotEnc() As Currency
Dim laSdoEncInt() As String, laTotEncInt() As Currency
Dim laSdoBil() As String, laTotBil() As Currency
Dim laProvis() As String, laTotPro() As Currency
Dim ldFecIni As Date, ldFecFin As Date
Dim lsOpeCod As String
Dim lnCntPag As Integer
Dim lsRTFEnc As String
Dim nPPD As Currency
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim nDias As Integer, lnTipo As Integer
Dim lnNroAgencias As Integer

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing
    
    Dim oCons As New NConstSistemas
    lsEmprLogo = oCons.LeeConstSistema(gConstSistNombreAbrevCMAC)
    Set oCons = Nothing
    
Dim oIni As ClasIni
   Set oIni = New ClasIni
   vsBaseComunes = oIni.BaseComunes
   vsBasePesonas = oIni.BasePersonas
   sLPT = "LPT1"
   BON = oImpresora.gPrnBoldON
   BOFF = oImpresora.gPrnBoldOFF
   CON = oImpresora.gPrnCondensadaON
   COFF = oImpresora.gPrnCondensadaOFF
   Set oIni = Nothing
End Sub

Public Sub Inicio(psInstCmac As String, pdFecsis As Date)
gsInstCmac = psInstCmac
gdFecSis = pdFecsis
End Sub

Public Function ImprimeResumenFlujoDiario(pdFecha As Date, psOpeCod As String) As String
    Dim HabilIng As String
    Dim OperaIng As String
    Dim SobraIng As String
    Dim OtrosIng As String
    Dim HabilEgr As String
    Dim OperaEgr As String
    Dim FaltaEgr As String
    Dim OtrosEgr As String
    Dim lnTotalSaldo As Currency
    Dim sTexto   As String
    Dim nTotal   As Currency
    Dim nImporte As Currency
    Dim nSaldoCaja As Currency
    Dim sCod       As String
Dim oOpe As New DOperacion
Dim oSdo As New NCtasaldo
Dim oEfe As New Defectivo
Dim sCtaCaja As String

On Error GoTo ImprimeResumenFlujoDiarioErr

    sTexto = ""
    nTotal = 0
    gsSimbolo = gcMN
    If Mid(psOpeCod, 3, 1) = 2 Then
        gsSimbolo = gcME
    End If
'    If Mid(psOpeCod, 3, 1) = "2" Then
'        HabilIng = "'402004'"
'        OperaIng = "'402202','402203','402207','400041'"
'        SobraIng = "'999999'"
'        OtrosIng = "'402011','402012','402330','602300','602800'"
'        HabilEgr = "'402003'"
'        OperaEgr = "'402201','402206','400042','402301','402302','402303'"
'        FaltaEgr = "'999999'"
'        OtrosEgr = "'402021','402022','402310','502301','602305','401204','402701','402703'"
'        gsSimbolo = gcME
'    Else
'        HabilIng = "'401004'"
'        OperaIng = "'401202','401203','401207','400042'"
'        SobraIng = "'999999'"
'        OtrosIng = "'401011','401012','401330','601300','601800','602300'"
'        HabilEgr = "'401003'"
'        OperaEgr = "'401201','401206','400041','401301','401302','401303'"
'        FaltaEgr = "'999999'"
'        OtrosEgr = "'401021','401022','401310','501301','601305','401204','601300','401701','601800'"
'        gsSimbolo = gcMN
'    End If
    Set rs = oOpe.CargaOpeCta(psOpeCod, , "0")
    If rs.EOF Then
        Err.Raise 50001, "NCajaGenImprimir: ImprimeDetalleFlujoDiario", "No se asignó Cuenta de Caja a Operación"
        Exit Function
    End If
    sCtaCaja = rs!cCtaContCod
    
    Linea sTexto, Cabecera("RESUMEN DIARIO DE FLUJO DE CAJA GENERAL", 0, IIf(Mid(psOpeCod, 3, 1) = "1", gcMN, gcME), 80, , , lsEmprLogo)
    Linea sTexto, BON & Space(62) & "FECHA : " & pdFecha & BOFF, 2

    nImporte = oSdo.GetCtaSaldo(sCtaCaja, Format(CDate(pdFecha) - 1, gsFormatoFecha), gsSimbolo = gcMN)
    nSaldoCaja = nImporte
    Linea sTexto, BON & Mid("1. SALDO DEL DIA ANTERIOR" & Space(55), 1, 55) & " " & gsSimbolo & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14) & BOFF
    Linea sTexto, BON & "2. INGRESOS DEL DIA" & BOFF
    
'
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, HabilIng)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   HABILITACIONES RECIBIDAS" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, OperaIng)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   INGRESOS POR OPERACIONES" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, SobraIng)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   SOBRANTES               " & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'
'    'nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, OtrosIng)
'    nImporte = GetOtrosIngEgres(sCtaCaja, pdFecha, OtrosIng)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   OTROS INGRESOS " & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'
    nTotal = 0
    Set rs = oEfe.Operacion_FlujoDiariaCaja(sCtaCaja, pdFecha, psOpeCod, True)
    Do While Not rs.EOF
        If gsSimbolo = gcMN Then
            nImporte = rs!nMovImporte
        Else
            nImporte = rs!nMovMEImporte
        End If
        nTotal = nTotal + nImporte
        Linea sTexto, Mid("   " & rs!cOpeDesc & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
        rs.MoveNext
    Loop
    nSaldoCaja = nSaldoCaja + nTotal
    Linea sTexto, Space(43) & "----------------"
    Linea sTexto, Space(23) & Right(Space(20) & "TOTAL INGRESOS  " & gsSimbolo, 20) & " " & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14), 2
    
    nTotal = 0
    Linea sTexto, BON & "3. EGRESOS DEL DIA" & BOFF
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, HabilEgr)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   HABILITACIONES ENTREGADAS" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, OperaEgr)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   EGRESOS POR OPERACIONES" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'    nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, FaltaEgr)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   FALTANTES" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'    'nImporte = GetMovOpeImporte(sCtaCaja, pdFecha, OtrosEgr)
'    nImporte = GetOtrosIngEgres(sCtaCaja, pdFecha, OtrosEgr, False)
'    nTotal = nTotal + nImporte
'    Linea sTexto, Mid("   OTROS EGRESOS" & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
'
'    Linea sTexto, Space(43) & "----------------"
'    Linea sTexto, Space(23) & Right(Space(20) & "TOTAL EGRESOS  " & gsSimbolo, 20) & " " & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14), 2
    
    Set rs = oEfe.Operacion_FlujoDiariaCaja(sCtaCaja, pdFecha, psOpeCod, False)
    Do While Not rs.EOF
        If gsSimbolo = gcMN Then
            nImporte = rs!nMovImporte
        Else
            nImporte = rs!nMovMEImporte
        End If
        nTotal = nTotal + nImporte
        Linea sTexto, Mid("   " & rs!cOpeDesc & Space(43), 1, 43) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
        rs.MoveNext
    Loop
    Linea sTexto, Space(43) & "----------------"
    Linea sTexto, Space(23) & Right(Space(20) & "TOTAL EGRESOS  " & gsSimbolo, 20) & " " & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14), 2
    nSaldoCaja = nSaldoCaja - nTotal
    Linea sTexto, BON & Mid("4. SALDO EFECTIVO DEL DIA" & Space(55), 1, 55) & " " & gsSimbolo & Right(Space(14) & Format(nSaldoCaja, gsFormatoNumeroView), 14) & BOFF

    Linea sTexto, BON & "5. DESCOMPOSICION DE EFECTIVO" & BOFF
    nTotal = 0
    Set rs = oEfe.GetBilletajes(Mid(psOpeCod, 3, 1))
    lnTotalSaldo = 0
    Do While Not rs.EOF
        sCod = Mid(rs!cEfectivoCod, 2, 1)
        Do While sCod = Mid(rs!cEfectivoCod, 2, 1)
            nImporte = oSdo.GetCtaEfectivoSaldo(sCtaCaja, Format(pdFecha, gsFormatoFecha), Mid(psOpeCod, 3, 1) = "1", rs!cEfectivoCod)
            Linea sTexto, "   " & Mid(rs!Descripcion & Space(40), 1, 40) & " " & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14)
            nTotal = nTotal + nImporte
            rs.MoveNext
            If rs.EOF Then
                Exit Do
            End If
        Loop
        Linea sTexto, Space(43) & "----------------"
        Linea sTexto, Space(23) & Right(Space(20) & "TOTAL " & IIf(sCod = "B", "BILLETES", "MONEDAS") & " " & gsSimbolo, 20) & " " & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14), 2
        lnTotalSaldo = lnTotalSaldo + nTotal
        nTotal = 0
    Loop
    Linea sTexto, Space(43) & "================"
    Linea sTexto, Space(23) & Right(Space(20) & "TOTAL EFECTIVO" & " " & gsSimbolo, 20) & " " & Right(Space(14) & Format(lnTotalSaldo, gsFormatoNumeroView), 14), 2

    'nSaldoCaja cambaido por lnTotalSaldo
    Dim oDocRec As NDocRec
    Set oDocRec = New NDocRec
    nTotal = oDocRec.GetImporteChequesCartera(Mid(psOpeCod, 3, 1), pdFecha)
    Set oDocRec = Nothing
    Linea sTexto, BON & Mid("6. TOTAL CHEQUES EN CARTERA" & Space(55), 1, 55) & " " & gsSimbolo & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14) & BOFF
    Linea sTexto, BON & Mid("7. TOTAL CHEQUES MAS EFECTIVO" & Space(55), 1, 55) & " " & gsSimbolo & Right(Space(14) & Format(nSaldoCaja + nTotal, gsFormatoNumeroView), 14) & BOFF
    
    RSClose rs
    ImprimeResumenFlujoDiario = sTexto
Exit Function
ImprimeResumenFlujoDiarioErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ImprimeDetalleFlujoDiario(pdFechaDel As Date, pdFechaAl As Date, psOpeCod As String) As String
    Dim nLin As Integer
    Dim P As Integer
    Dim sDia As String
    Dim oOpe  As New DOperacion
    Dim oCont As New NContAsientos
    Dim oSdo  As New NCtasaldo
    Dim sCtaCaja As String
    Dim sTexto   As String
    Dim nTotal   As Currency, nTotalH As Currency
    Dim nImporte As Currency, nSaldoCaja As Currency
    Dim oCon As DConecta
    Dim sSQL As String
    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    On Error GoTo ImprimeDetalleFlujoDiarioErr
    
    nLin = 66
    P = 0
    nTotal = 0: nTotalH = 0
    sTexto = ""

    If Mid(psOpeCod, 3, 1) = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    Set rs = oOpe.CargaOpeCta(psOpeCod)
    If rs.EOF Then
        Err.Raise 50001, "NCajaGenImprimir: ImprimeDetalleFlujoDiario", "No se asignó Cuenta de Caja a Operación"
        Exit Function
    End If
    sCtaCaja = rs!cCtaContCod
    Set rs = oCont.GetMayorCuenta(sCtaCaja, Format(pdFechaDel, gsFormatoMovFecha), Format(pdFechaAl, gsFormatoMovFecha), , , gsSimbolo = gcME, False)
    If rs.EOF Then
        Err.Raise 50001, "NCajaGenImprimir: ImprimeDetalleFlujoDiario", "No se registraron Movimientos en Rango de Fechas especificadas"
        Exit Function
    End If
    Linea sTexto, DetalleFlujoDiarioCabecera(nLin, P, Format(pdFechaDel, gsFormatoFechaView), Format(pdFechaAl, gsFormatoFechaView))
    
    nImporte = oSdo.GetCtaSaldo(sCtaCaja, Format(pdFechaDel - 1, gsFormatoFecha), gsSimbolo = gcMN)
    nSaldoCaja = nImporte
    Linea sTexto, BON & CON & Mid("SALDO DEL DIA ANTERIOR" & Space(117), 1, 117) & " " & gsSimbolo & Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14) & COFF & BOFF, , nLin
    sDia = Mid(rs!cMovNro, 1, 8)
    
    Do While Not rs.EOF
        Linea sTexto, DetalleFlujoDiarioCabecera(nLin, P, Format(pdFechaDel, gsFormatoFechaView), Format(pdFechaAl, gsFormatoFechaView)), 0
        If gsSimbolo = gcMN Then
            nImporte = rs!nDebe - rs!nHaber
        Else
            nImporte = rs!nDebeME - rs!nHaberME
        End If
        If nImporte > 0 Then
            nTotal = nTotal + nImporte
        Else
            nTotalH = nTotalH + (nImporte * -1)
        End If
        nSaldoCaja = nSaldoCaja + nImporte
        If sDia <> Mid(rs!cMovNro, 1, 8) Then
            sDia = Mid(rs!cMovNro, 1, 8)
            Linea sTexto, ""
            nLin = nLin + 1
        End If
        
        sSQL = "SELECT d.cDocAbrev, md.cDocNro, md.dDocFecha  " _
            & " FROM    MovDoc md " _
            & "         LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
            & "         JOIN MOV M on M.nMovNro= MD.nMovNro  " _
            & "WHERE    M.cMovNro = '" & rs!cMovNro & "' " _
            & " union " _
            & " SELECT  d.cDocAbrev, md.cDocNro, md.dDocFecha   " _
            & " FROM    movref  mr " _
            & "         join movdoc md on md.nmovnro = mr.nMovNroRef " _
            & "         left join documento d on d.nDocTpo = md.nDocTpo " _
            & "         JOIN MOV M ON M.NMOVNRO = MR.nMovNroRef join mov m1 on m1.nmovnro = mr.nmovnro  " _
            & " WHERE   m1.cMovNro = '" & rs!cMovNro & "' AND M.NMOVFLAG=0 "
        
        Dim rsDoc As ADODB.Recordset
        Dim sDoc As String
        Set rsDoc = oCon.CargaRecordSet(sSQL)
        sDoc = ""
        If Not rsDoc.EOF Then
           Do While Not rsDoc.EOF
              sDoc = sDoc & ImpreFormat(rsDoc!cDocAbrev & "-" & rsDoc!cDocNro, 25) & rsDoc!dDocFecha & Space(5)
              rsDoc.MoveNext
           Loop
        End If
        
        Linea sTexto, CON & " " & ImpreMovNro(rs!cMovNro) & ImpreFormat(Trim(sDoc), 36) & " " & ImpreFormat(Replace(Replace(Replace(rs!cMovdesc, Chr(13), ""), Chr(10), ""), Chr(13) + Chr(10), ""), 30) & " " & IIf(nImporte > 0, Right(Space(14) & Format(nImporte, gsFormatoNumeroView), 14), Space(14)) & " " & IIf(nImporte < 0, Right(Space(14) & Format(nImporte * -1, gsFormatoNumeroView), 14), Space(14)) & " " & Right(Space(14) & Format(nSaldoCaja, gsFormatoNumeroView), 14) & COFF, , nLin
        rs.MoveNext
    Loop
    Linea sTexto, CON & String(148 - 1, "-")
    Linea sTexto, BON & Space(686) & "TOTALES  " & gsSimbolo & " " & Right(Space(14) & Format(nTotal, gsFormatoNumeroView), 14) & " " & Right(Space(14) & Format(nTotalH, gsFormatoNumeroView), 14) & BOFF
    Linea sTexto, String(148 - 1, "=") & COFF
    ImprimeDetalleFlujoDiario = sTexto
    
    oCon.CierraConexion
    Set oCon = Nothing
Exit Function
ImprimeDetalleFlujoDiarioErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function DetalleFlujoDiarioCabecera(nLin As Integer, P As Integer, dFecha1 As String, dFecha2 As String) As String
    Dim sCab As String
    If nLin > 60 Then
        Linea sCab, COFF & Cabecera("REPORTE DETALLADO DE FLUJO DE CAJA GENERAL", P, gsSimbolo, 80, , , lsEmprLogo)
        Linea sCab, Centra("( DEL " & dFecha1 & " AL " & dFecha2 & " )", gnColPage), 2
        Linea sCab, CON & String(148, "-")
        Linea sCab, " Número de                           DOCUMENTOS                             CONCEPTO                     MOVIMIENTO              SALDOS"
        Linea sCab, " Movimiento                                                                                         DEBE            HABER              "
        Linea sCab, String(148 - 1, "-") & COFF
        nLin = 12
        DetalleFlujoDiarioCabecera = sCab
    End If
End Function


Public Function ImprimeCartasFianzaSalida(pdFecha As Date, pnMoneda As Integer, pnTipCambio As Currency, pnLinPage As Integer, psCtaCod As String) As String
Dim rs As ADODB.Recordset
Dim nLin As Integer
Dim P    As Integer
Dim sTit As String, sTitBarra As String
Dim sTexto As String
Dim nTotDif As Currency
Dim oCarta  As New nCajaGeneral
On Error GoTo ImprimeCartasFianzaSalidaErr

sTitBarra = "RESUMEN DE GARANTIAS"
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, sTitBarra, "", "Obteniendo datos", vbBlue)
Set rs = oCarta.GetDatosCartaFianza(psCtaCod)
RaiseEvent BarraClose
If rs.EOF Then
   Err.Raise "50001", "NCajaGenImprimi: ImprimeCartasFianzaSalida", "No existen Datos para generar Cartas Fianza Vigentes"
Else
   nLin = pnLinPage
   P = 0
   sTit = "R E P O R T E   D E   C A R T A S   F I A N Z A"
   RaiseEvent BarraShow(rs.RecordCount)
   Do While Not rs.EOF
      If nLin > pnLinPage - 6 Then
         Linea sTexto, Cabecera(sTit, P, IIf(Mid(psCtaCod, 3, 1) = gMonedaNacional, gcMN, gcME), 80, , , lsEmprLogo), 0
         Linea sTexto, CON & " Nro.Carta           Fec.Emis.   Persona                        Concepto                       Entidad Financiera                 Monto   Fecha Ing. Fecha Venc."
         Linea sTexto, " -------------------------------------------------------------------------------------------------------------------------------------------------------------------" & COFF
         nLin = 5
      End If
                         
      Linea sTexto, CON & " " & Justifica(rs!cDocNro, 19) & " " & Format(rs!dDocFecha, gsFormatoFechaView) & "  " & _
                    Justifica(rs!cpersNombre, 30) & " " & Left(rs!Descripcion, 40) & " " & Justifica(rs!cBanco, 30) & PrnVal(rs!nMovImporte, 12, 2) & " " & rs!dFecIng & " " & Format(rs!FechaVenc, "dd/MM/YYYY") & COFF
      nTotDif = nTotDif + rs!nMovImporte
      nLin = nLin + 1
      rs.MoveNext
   Loop
   Linea sTexto, CON & " ---------------------------------------------------------------------------------------------------------------------------------------------------------------" & COFF
   Linea sTexto, CON & Space(114) & "TOTAL : " & PrnVal(nTotDif, 15, 2)
   ImprimeCartasFianzaSalida = sTexto
   RaiseEvent BarraClose
End If
Exit Function
ImprimeCartasFianzaSalidaErr:
    Call RaiseError(MyUnhandledError, "NCajaGenImprimir:ImprimeCartasFianzaSalida Method")
End Function


Public Function ExcelCartasFianzaSalida(pdFecha As Date, pnMoneda As Integer, pnTipCambio As Currency, pnLinPage As Integer, psCtaCod As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, gsNomCmac As String, gsNomAge As String)
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lnFil As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    Dim oCarta  As New nCajaGeneral
    nTotAge = 0
    Set rs = oCarta.GetDatosCartaFianza(psCtaCod)
    If rs.EOF Then
       Err.Raise "50001", "NCajaGenImprimi: ImprimeCartasFianzaSalida", "No existen Datos para generar Cartas Fianza Vigentes"
        ExcelCartasFianzaSalida = ""
    Else
        ExcelCartasFianzaSalida = "Reporte con datos"
        ExcelAddHoja "Carta Fianza", xlLibro, xlHoja1
        xlHoja1.PageSetup.Orientation = xlLandscape
        xlHoja1.PageSetup.Zoom = False
        xlHoja1.PageSetup.FitToPagesTall = 1
        xlHoja1.PageSetup.FitToPagesWide = 1
        'xlHoja1.PageSetup.CenterVertically = True
        
      
        xlHoja1.Cells(1, 1) = Trim(gsNomCmac)
        xlHoja1.Cells(1, 9) = Trim(gsNomAge)
        xlHoja1.Cells(2, 1) = "R E P O R T E   D E   C A R T A S   F I A N Z A"
        xlHoja1.Cells(3, 1) = IIf(Mid(psCtaCod, 3, 1) = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
        
        xlHoja1.Range("A2:I2").MergeCells = True
        'xlHoja1.Range("A3:I3").MergeCells = True
        
        lnFil = 5
        
        lnFil = lnFil + 1
        xlHoja1.Cells(lnFil, 1) = "Nro. Carta"
        xlHoja1.Cells(lnFil, 2) = "Fec. Emision"
        xlHoja1.Cells(lnFil, 3) = "Persona"
        xlHoja1.Cells(lnFil, 4) = "Concepto"
        xlHoja1.Cells(lnFil, 5) = "Entidad Financiera"
        xlHoja1.Cells(lnFil, 6) = "Monto"
        xlHoja1.Cells(lnFil, 7) = "Fec. Ing."
        xlHoja1.Cells(lnFil, 8) = "Fec. Venc."
        xlHoja1.Cells(lnFil, 9) = "Movimiento"
         
        'Tipos de Letra formateados
        'xlHoja1.Range("F1:F1").Font.Size = 8
        'xlHoja1.Range("A25:F32").Font.Size = 8
        'xlHoja1.Range("A4:A4").Font.Size = 11
        xlHoja1.Range("A1:I6").Font.Bold = True
        
        'xlHoja1.Range("A4:A4").Font.Underline = True
        'xlHoja1.Range("A29:A29").Font.Underline = True
         
        Do While Not rs.EOF
            lnFil = lnFil + 1
            xlHoja1.Cells(lnFil, 1) = rs!cDocNro
            xlHoja1.Cells(lnFil, 2) = "'" & Format(rs!dDocFecha, gsFormatoFechaView)
            xlHoja1.Cells(lnFil, 3) = rs!cpersNombre
            xlHoja1.Cells(lnFil, 4) = Left(rs!Descripcion, 40)
            xlHoja1.Cells(lnFil, 5) = rs!cBanco
            xlHoja1.Cells(lnFil, 6) = Format(rs!nMovImporte, "#,###,###.00")
            xlHoja1.Cells(lnFil, 7) = "'" & Format(rs!dFecIng, "dd/MM/YYYY")
            xlHoja1.Cells(lnFil, 8) = "'" & Format(rs!FechaVenc, "dd/MM/YYYY")
            xlHoja1.Cells(lnFil, 9) = rs!cMovNro
            nTotAge = nTotAge + rs!nMovImporte
            rs.MoveNext
        Loop
    End If
    xlHoja1.Cells(lnFil + 1, 5) = "TOTAL :"
    xlHoja1.Cells(lnFil + 1, 6) = nTotAge
    Set rs = Nothing
     
    'Tipo de letra de todo el reporte
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    
    'Alineacion
    xlHoja1.Range("A2:I6").HorizontalAlignment = xlCenter
    xlHoja1.Range("B7:B" & Trim(Str(lnFil))).HorizontalAlignment = xlCenter
    xlHoja1.Range("G7:G" & Trim(Str(lnFil))).HorizontalAlignment = xlCenter
    xlHoja1.Range("H7:H" & Trim(Str(lnFil))).HorizontalAlignment = xlCenter
    xlHoja1.Range("A3:I3").MergeCells = True
    
    'Cuadro
    xlHoja1.Range("A6:I6").Borders.LineStyle = xlContinuous
    
    xlHoja1.Range("A6:I" & Trim(Str(lnFil))).BorderAround xlContinuous
    
    xlHoja1.Range("A6:I" & Trim(Str(lnFil))).Borders(xlInsideVertical).LineStyle = xlContinuous

End Function

Public Function ImprimeCartasFianza(pdFechaDel As Date, pdFechaAl As Date, pbIngreso As Boolean, psOpeCod As String, pdFecsis As Date, ByRef psMensaje As String) As String
Dim lnPaginas  As Integer
Dim lnLineas   As Integer
Dim Total      As Currency
Dim lsCtaCod   As String
Dim lsImpre    As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oOpe As New DOperacion
On Error GoTo CartaErr
lsCtaCod = oOpe.EmiteOpeCta(psOpeCod, "D")
Set oOpe = Nothing


Dim oCaja As nCajaGeneral
Set oCaja = New nCajaGeneral
Set rs = oCaja.GetDatosCartaFianza(lsCtaCod, IIf(pbIngreso, 1, 2), Format(pdFechaDel, gsFormatoMovFecha), Format(pdFechaAl, gsFormatoMovFecha))
If rs.EOF Then
    RSClose rs
    Set oCaja = Nothing
    
    psMensaje = "No existen Datos en Rango de Fechas para Imprimir"
    Exit Function
    'Err.Raise 50001, "NCajaGenImprimir: ImprimeCartaFianza", "No existen Datos en Rango de Fechas para Imprimir"
End If
lnPaginas = 0
lnLineas = 66
Do While Not rs.EOF
   Linea lsImpre, ImprimeCartasFianzaCabe(pdFechaDel, pdFechaAl, pbIngreso, lnLineas, lnPaginas, Val(Mid(psOpeCod, 3, 1)), pdFecsis), 0
   If pbIngreso = True Then
      Linea lsImpre, CON & ImpreFormat(Trim(rs!cDocNro), 22, 0) & ImpreFormat(Trim(PstaNombre(rs!cpersNombre)), 25) & ImpreFormat(Trim(Replace(rs!Descripcion, Chr(13), "")), 25) & ImpreFormat(rs!cBanco, 18) & PrnVal(rs!nMovImporte, 12, 2) & _
               ImpreFormat(GetFechaMov(rs!cMovNro, True), 12) & ImpreFormat(Format(rs!dDocFecha, gsFormatoFechaView), 12) & ImpreFormat(Format(rs!FechaVenc, gsFormatoFechaView), 12) & COFF
   Else
      Linea lsImpre, CON & ImpreFormat(Trim(rs!cDocNro), 22, 0) & ImpreFormat(Trim(PstaNombre(rs!cpersNombre)), 25) & ImpreFormat(Trim(Replace(rs!Descripcion, Chr(13), "")), 25) & ImpreFormat(rs!cBanco, 18) & PrnVal(rs!nMovImporte, 12, 2) & _
               ImpreFormat(GetFechaMov(rs!cMovNro, True), 12) & ImpreFormat(Format(rs!FechaVenc, gsFormatoFechaView), 12) & ImpreFormat(GetFechaMov(rs!cMovNroSalida, True), 12) & COFF
   End If
   Total = Total + rs!nMovImporte
   lnLineas = lnLineas + 1
   rs.MoveNext
Loop
RSClose rs
Linea lsImpre, CON & String(150, "-")
Linea lsImpre, BON & ImpreFormat("TOTAL :", 12, 95) & PrnVal(Total, 12, 2) & BOFF & COFF
ImprimeCartasFianza = lsImpre
Exit Function
CartaErr:
    Err.Raise Err.Number, "ImprimeCartasFianza", Err.Description
End Function

Private Function ImprimeCartasFianzaCabe(pdFechaDel As Date, pdFechaAl As Date, pbIngreso As Boolean, ByRef pnLineas As Integer, ByRef pnNumPagina As Integer, pnMoneda As Integer, pdFecsis As Date)
Dim lsTitulo As String
Dim lsTexto  As String
If pbIngreso = True Then
    lsTitulo = "REPORTE DE CARTAS FIANZAS INGRESADAS"
Else
    lsTitulo = "REPORTE DE CARTAS FIANZAS ENTREGADAS"
End If
lsTitulo = lsTitulo & IIf(pnMoneda = 1, " MONEDA NACIONAL", " MONEDA EXTRANJERA")
lsTexto = ""
If pnLineas > 60 Then
   If pnNumPagina > 0 Then
      pnNumPagina = pnNumPagina + 1
      lsTexto = oImpresora.gPrnSaltoPagina
   End If
   Linea lsTexto, PrnSet("MI", 5) & Cabecera(lsTitulo, pnNumPagina, "", 80, _
            Centra(" DESDE " & pdFechaDel & " AL " & pdFechaAl), Format(pdFecsis, gsFormatoFechaView), lsEmprLogo)
    
   Linea lsTexto, CON & String(150, "-")
   If pbIngreso = True Then
       Linea lsTexto, ImpreFormat("No CARTA FIANZA", 22, 0) & ImpreFormat("PROVEEDOR", 25) & ImpreFormat("DESCRIPCION", 25) & ImpreFormat("ENTIDAD FINANCIERA", 23) & ImpreFormat("MONTO", 6) & ImpreFormat("FECHA ING.", 12) & ImpreFormat("FECHA EMIS.", 12) & ImpreFormat("FECHA VENC.", 12)
   Else
       Linea lsTexto, ImpreFormat("No CARTA FIANZA", 22, 0) & ImpreFormat("PROVEEDOR", 25) & ImpreFormat("DESCRIPCION", 25) & ImpreFormat("ENTIDAD FINANCIERA", 23) & ImpreFormat("MONTO", 6) & ImpreFormat("FECHA ING.", 12) & ImpreFormat("FECHA VENC.", 12) & ImpreFormat("FEC.ENTREG.", 12)
   End If
   Linea lsTexto, String(150, "-") & COFF
   pnLineas = 6
   
End If
ImprimeCartasFianzaCabe = lsTexto
End Function

Public Function ImprimePagoProveedoresEntrega(rsImpre As ADODB.Recordset, pdFecsis As Date, pnLinPage As Integer) As String
Dim lsImpre As String
Dim I As Integer
Dim lnLin As Integer
Dim lnPag As Integer
Dim lnTotal As Currency

If Not rsImpre Is Nothing Then
   lnLin = pnLinPage
   lnPag = 0
   lsImpre = PrnSet("MI", 3)
   lnTotal = 0
   rsImpre.MoveFirst
   RaiseEvent BarraShow(rsImpre.RecordCount)
   Do While Not rsImpre.EOF
      If lnLin > pnLinPage - 6 Then
         Linea lsImpre, Cabecera("REPORTE DE DOCUMENTOS DE PAGO EMITIDAS NO ENTREGADAS", lnPag, "", 80, "", Format(pdFecsis, gsFormatoFechaView), lsEmprLogo), 0
         lnLin = 4
         Linea lsImpre, CON & String(115, "-")
         Linea lsImpre, "Itm  Comprobante     Emisión     Proveedor                           Concepto                            Importe "
         Linea lsImpre, String(115, "-") & COFF
         lnLin = lnLin + 3
      End If
      lnTotal = lnTotal + rsImpre!Importe
      Linea lsImpre, CON & ImpreFormat(rsImpre!Ord, 3, 0), 0
      Linea lsImpre, ImpreFormat(rsImpre!Comprobante, 14) & " ", 0
      Linea lsImpre, rsImpre!Emisión & " ", 0
      Linea lsImpre, ImpreFormat(rsImpre!Proveedor, 32) & " ", 0
      Linea lsImpre, ImpreFormat(rsImpre!cMovdesc, 27) & " ", 0
      Linea lsImpre, PrnVal(rsImpre!Importe, 15, 2) & COFF
      lnLin = lnLin + 1
      RaiseEvent BarraProgress(rsImpre.Bookmark, "REPORTE DE DOCUMENTOS DE PAGO", "", "Generando Reporte...", vbBlue)
      rsImpre.MoveNext
   Loop
   Linea lsImpre, CON & String(115, "=")
   Linea lsImpre, ImpreFormat("", 82) & ImpreFormat("Total :", 10) & ImpreFormat(lnTotal, 15, 2, True) & COFF
   RaiseEvent BarraClose
End If
ImprimePagoProveedoresEntrega = lsImpre
End Function

''''''''''''''''''''''''Public Function ImprimeConsolidaSdoEnc(pnBitCentral_ As Boolean, psOpeCod As String, psFecIni As String, psFecFin As String, pnTipo As Integer, _
''''''''''''''''''''''''       xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
''''''''''''''''''''''''Dim i As Integer, J As Integer, K As Integer
''''''''''''''''''''''''
''''''''''''''''''''''''Dim dFecha As Date
''''''''''''''''''''''''Dim lsCodCtaCaja As String
''''''''''''''''''''''''Dim lsCodCtaBanc As String
''''''''''''''''''''''''Dim lsCodCtaCAge As String
''''''''''''''''''''''''Dim lsMoneda     As String
''''''''''''''''''''''''Dim lnContAge      As Integer
''''''''''''''''''''''''
''''''''''''''''''''''''Dim oDOpe As New DOperacion
''''''''''''''''''''''''Dim oSdo  As New NCtasaldo
''''''''''''''''''''''''Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''Dim lnDiaMes As Integer
''''''''''''''''''''''''Dim dTempo As Date
''''''''''''''''''''''''Dim nTotal As Double
''''''''''''''''''''''''
''''''''''''''''''''''''lsOpeCod = psOpeCod
''''''''''''''''''''''''lsMoneda = Mid(psOpeCod, 3, 1)
''''''''''''''''''''''''lnTipo = pnTipo
''''''''''''''''''''''''RaiseEvent BarraShow(1)
''''''''''''''''''''''''RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "", "Obteniendo datos preliminares..", vbBlue)
''''''''''''''''''''''''
''''''''''''''''''''''''lsCodCtaCAge = oDOpe.EmiteOpeCta(psOpeCod, "H", "0")
''''''''''''''''''''''''lsCodCtaCaja = oDOpe.EmiteOpeCta(psOpeCod, "H", "1")
''''''''''''''''''''''''lsCodCtaBanc = oDOpe.EmiteOpeCta(psOpeCod, "H", "2")
''''''''''''''''''''''''
''''''''''''''''''''''''Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''lnNroAgencias = rs.RecordCount
''''''''''''''''''''''''RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "", "Obteniendo datos preliminares..", vbBlue)
''''''''''''''''''''''''RaiseEvent BarraClose
''''''''''''''''''''''''
''''''''''''''''''''''''ldFecIni = CDate(psFecIni)
''''''''''''''''''''''''ldFecFin = CDate(psFecFin)
''''''''''''''''''''''''nDias = ldFecFin - ldFecIni + 1
''''''''''''''''''''''''
''''''''''''''''''''''''' Inicializamos variables
''''''''''''''''''''''''ReDim laSdoEnc(1 To lnNroAgencias, 1 To 3, 1 To nDias)  'Saldos de Encajes
''''''''''''''''''''''''ReDim laTotEnc(1 To lnNroAgencias, 1 To 3)              'Totales Enc por Agencia
''''''''''''''''''''''''
''''''''''''''''''''''''ReDim laSdoEncInt(1 To lnNroAgencias + 1, 1 To 3, 1 To nDias) 'Saldos de Encajes para interes
''''''''''''''''''''''''ReDim laTotEncInt(1 To lnNroAgencias + 1, 1 To 3)             'Totales Enc por Agencia para interes
''''''''''''''''''''''''
''''''''''''''''''''''''ReDim laSdoBil(1 To lnNroAgencias + 2, 1 To nDias)      'Saldo de Billetes por Agencia
''''''''''''''''''''''''ReDim laTotBil(1 To lnNroAgencias + 2)                  'Totales Billete por Agencia
''''''''''''''''''''''''ReDim laProvis(1 To 8, 1 To nDias)                       'Para Planilla de Provision
''''''''''''''''''''''''ReDim laTotPro(1 To 8)                                   'Totales por Concepto de Provision
''''''''''''''''''''''''For K = 1 To lnNroAgencias
''''''''''''''''''''''''   For i = 1 To nDias
''''''''''''''''''''''''      For J = 1 To 3
''''''''''''''''''''''''         laSdoEnc(K, J, i) = "0"
''''''''''''''''''''''''         laSdoEncInt(K, J, i) = "0"
''''''''''''''''''''''''      Next J
''''''''''''''''''''''''   Next i
''''''''''''''''''''''''   laTotBil(K) = 0
''''''''''''''''''''''''Next K
''''''''''''''''''''''''laTotBil(lnNroAgencias + 1) = 0
''''''''''''''''''''''''laTotBil(lnNroAgencias + 2) = 0
''''''''''''''''''''''''
''''''''''''''''''''''''For K = 1 To lnNroAgencias
''''''''''''''''''''''''   For i = 1 To 3
''''''''''''''''''''''''      laTotEnc(K, i) = 0
''''''''''''''''''''''''      laTotEncInt(K, i) = 0
''''''''''''''''''''''''   Next
''''''''''''''''''''''''   For i = 1 To nDias
''''''''''''''''''''''''      laSdoEnc(K, 1, i) = CDate(psFecIni) + i - 1
''''''''''''''''''''''''      laSdoBil(K, i) = "0"
''''''''''''''''''''''''   Next
''''''''''''''''''''''''Next
''''''''''''''''''''''''For i = 1 To nDias
''''''''''''''''''''''''   For J = 1 To 8
''''''''''''''''''''''''      laProvis(J, i) = "0"
''''''''''''''''''''''''      laTotPro(J) = 0
''''''''''''''''''''''''   Next
''''''''''''''''''''''''Next
''''''''''''''''''''''''If lnTipo = 1 Then
''''''''''''''''''''''''
''''''''''''''''''''''''   'Cuenta Contable de efectivo para BilletajeCaja
''''''''''''''''''''''''   CalculaBilleteCaja lsMoneda, lsCodCtaCaja, ldFecIni, ldFecFin
''''''''''''''''''''''''
''''''''''''''''''''''''   'Cuenta Contable de efectivo para Cuenta de BCR y objeto para cuenta de BCR
''''''''''''''''''''''''   CalculaSaldoBCR lsMoneda, lsCodCtaBanc, ldFecIni, ldFecFin
''''''''''''''''''''''''
''''''''''''''''''''''''End If
''''''''''''''''''''''''lnContAge = 0
''''''''''''''''''''''''Do While Not rs.EOF
''''''''''''''''''''''''   RaiseEvent BarraShow(1)
''''''''''''''''''''''''   RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", Trim(rs!Descripcion), "Procesando..", vbBlue)
''''''''''''''''''''''''   lnContAge = lnContAge + 1
''''''''''''''''''''''''   If rs!Nivel = 2 Then
''''''''''''''''''''''''
''''''''''''''''''''''''      CalculaSdoEnc pnBitCentral_, lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
''''''''''''''''''''''''
''''''''''''''''''''''''      If lnTipo = 1 Then
''''''''''''''''''''''''
''''''''''''''''''''''''         CalculaBilleteAge lsMoneda, rs!Codigo, lsCodCtaCAge, ldFecIni, ldFecFin, lnContAge
''''''''''''''''''''''''
''''''''''''''''''''''''      End If
''''''''''''''''''''''''      'CalculaSdoEncInteres lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
''''''''''''''''''''''''      CalculaSdoEncObligaciones lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
''''''''''''''''''''''''   End If
''''''''''''''''''''''''   RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", Trim(rs!Descripcion), "Procesando..", vbBlue)
''''''''''''''''''''''''   rs.MoveNext
''''''''''''''''''''''''Loop
''''''''''''''''''''''''
'''''''''''''''''''''''''Calculo Obligaciones Inmediatas Para Caja
''''''''''''''''''''''''
''''''''''''''''''''''''dTempo = ldFecIni
''''''''''''''''''''''''Do While Format(ldFecFin, "YYYYmmdd") >= Format(dTempo, "YYYYmmdd")
''''''''''''''''''''''''    lnDiaMes = dTempo - ldFecIni + 1
''''''''''''''''''''''''    nTotal = 0
''''''''''''''''''''''''    nTotal = SaldoObligInmediatas(1, "76" & lsMoneda & "201", dTempo, ldFecFin, "99")
''''''''''''''''''''''''    laSdoEncInt(UBound(laSdoEncInt, 1), 2, lnDiaMes) = nTotal
''''''''''''''''''''''''    laTotEncInt(UBound(laSdoEncInt, 1), 1) = Val(laTotEncInt(UBound(laSdoEncInt, 1), 1)) + laSdoEncInt(UBound(laSdoEncInt, 1), 2, lnDiaMes)
''''''''''''''''''''''''    dTempo = dTempo + 1
''''''''''''''''''''''''Loop
''''''''''''''''''''''''
''''''''''''''''''''''''If lsMoneda = "1" Then
''''''''''''''''''''''''   RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "CALCULO DE CAJA PROMEDIO", "Procesando..", vbBlue)
''''''''''''''''''''''''
''''''''''''''''''''''''   'En soles es el promedio del mes pasado
''''''''''''''''''''''''   CalculaBilleteTotCajaPromedio lsMoneda, lsCodCtaCaja, ldFecIni
''''''''''''''''''''''''
''''''''''''''''''''''''   RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "CALCULO DE CAJA PROMEDIO", "Procesando..", vbBlue)
''''''''''''''''''''''''End If
''''''''''''''''''''''''
''''''''''''''''''''''''Dim lsImpre As String
''''''''''''''''''''''''
''''''''''''''''''''''''If lnTipo = 1 Then
''''''''''''''''''''''''    ImprimePlanillaAhorroPF 2, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
''''''''''''''''''''''''
''''''''''''''''''''''''    ImprimePlanillaAhorroPF 3, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
''''''''''''''''''''''''
''''''''''''''''''''''''    ImprimePlanillaObligaciones 4, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
''''''''''''''''''''''''
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''        ImprimePlanillaBilletaje lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
''''''''''''''''''''''''
''''''''''''''''''''''''        CalculaTotProvis
''''''''''''''''''''''''        ImprimePlanillaProvision lsMoneda, Format(ldFecFin, gsFormatoFechaView), xlAplicacion, xlLibro, xlHoja1, psOpeCod
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    ImprimeConsolidaSdoEnc = ""
''''''''''''''''''''''''
'''''''''''''''''''''''''Quitado
'''''''''''''''''''''''''   lsImpre = CON & PrtSdoEnc(2, lsMoneda, xlAplicacion, xlLibro, xlHoja1) & PrtSdoEnc(3, lsMoneda, xlAplicacion, xlLibro, xlHoja1) & PrtSdoEncInt(4, lsMoneda, xlAplicacion, xlLibro, xlHoja1)
'''''''''''''''''''''''''   lsImpre = lsImpre & PrtSdoBillete(lsMoneda, xlAplicacion, xlLibro, xlHoja1)
'''''''''''''''''''''''''   CalculaTotProvis
''''''''''''''''''''''''
'''''''''''''''''''''''''   lsImpre = lsImpre & PrtPlanillaProvision(lsMoneda, Format(ldFecFin, gsFormatoFechaView), xlAplicacion, xlLibro, xlHoja1) & COFF
'''''''''''''''''''''''''Fin Quitado
''''''''''''''''''''''''
''''''''''''''''''''''''End If
''''''''''''''''''''''''RaiseEvent BarraClose
'''''''''''''''''''''''''ImprimeConsolidaSdoEnc = lsImpre
''''''''''''''''''''''''
''''''''''''''''''''''''End Function
''''''''''''''''''''''''
'''''''''''''''''''''''''Inicio Agregado
''''''''''''''''''''''''Private Sub ImprimePlanillaAhorroPF(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''    Dim cNameHoja As String
''''''''''''''''''''''''
''''''''''''''''''''''''    If nTipo = 2 Then
''''''''''''''''''''''''        cNameHoja = "AhorroC_Capital" 'ok
''''''''''''''''''''''''    ElseIf nTipo = 3 Then
''''''''''''''''''''''''        cNameHoja = "PlazoF_Capital" 'ok
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    ExcelAddHoja cNameHoja, xlLibro, xlHoja1
''''''''''''''''''''''''
''''''''''''''''''''''''    'ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = UBound(laSdoEnc, 1) * 20   '  180
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''      lsTitRp1 = "E F E C T O S   D E   E N C A J E (CAPITAL)"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    If nTipo = 2 Then
''''''''''''''''''''''''      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!Codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 1) = "TOTAL"
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoEnc, 3) To UBound(laSdoEnc, 3)
''''''''''''''''''''''''
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = 0
''''''''''''''''''''''''        For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
''''''''''''''''''''''''
''''''''''''''''''''''''            xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoEnc(K, nTipo, i))
''''''''''''''''''''''''
''''''''''''''''''''''''            nTotAge = nTotAge + nVal(laSdoEnc(K, nTipo, i))
''''''''''''''''''''''''        Next
''''''''''''''''''''''''
''''''''''''''''''''''''        laProvis(1, i) = laSdoEnc(1, 1, i)
''''''''''''''''''''''''        laProvis(nTipo, i) = nTotAge
''''''''''''''''''''''''        laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''        'Totales por Dia
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''    For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
''''''''''''''''''''''''        'Totales Por Agencia
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotEnc(K, nTipo)
''''''''''''''''''''''''        nTotAge = nTotAge + laTotEnc(K, nTipo)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells.Select
''''''''''''''''''''''''    xlHoja1.Cells.Font.Name = "Arial"
''''''''''''''''''''''''    xlHoja1.Cells.Font.Size = 9
''''''''''''''''''''''''    xlHoja1.Cells.EntireColumn.AutoFit
''''''''''''''''''''''''
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''Private Sub ImprimePlanillaObligaciones(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''    Dim cNameHoja As String
''''''''''''''''''''''''
''''''''''''''''''''''''    If nTipo = 4 Then
''''''''''''''''''''''''        cNameHoja = "Oblig_Inmediatas" ' ok
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja cNameHoja, xlLibro, xlHoja1
''''''''''''''''''''''''
''''''''''''''''''''''''    'ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = UBound(laSdoEncInt, 1) * 20   '  180
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''      lsTitRp1 = "E F E C T O S   D E   E N C A J E (INTERES)"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    If nTipo = 2 Then
''''''''''''''''''''''''      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    If nTipo = 4 Then
''''''''''''''''''''''''        lsTitRp1 = "O B L I G A C I O N E S   I N M E D I A T A S "
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!Codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "CAJA GENERAL"
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 2) = "TOTAL"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "TOTAL"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoEncInt, 3) To UBound(laSdoEncInt, 3)
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = 0
''''''''''''''''''''''''        For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
''''''''''''''''''''''''
''''''''''''''''''''''''            xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoEncInt(K, 2, i))
''''''''''''''''''''''''
''''''''''''''''''''''''            nTotAge = nTotAge + nVal(laSdoEncInt(K, 2, i))
''''''''''''''''''''''''        Next
''''''''''''''''''''''''
''''''''''''''''''''''''        'adiciona la provision para el total en la planilla provisional de encaje
''''''''''''''''''''''''        laProvis(nTipo, i) = nTotAge
''''''''''''''''''''''''        laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''        'Totales por Dia
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''
''''''''''''''''''''''''    For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
''''''''''''''''''''''''        'Totales Por Agencia
''''''''''''''''''''''''
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotEncInt(K, 1)
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = nTotAge + laTotEncInt(K, 1)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells.Select
''''''''''''''''''''''''    xlHoja1.Cells.Font.Name = "Arial"
''''''''''''''''''''''''    xlHoja1.Cells.Font.Size = 9
''''''''''''''''''''''''    xlHoja1.Cells.EntireColumn.AutoFit
''''''''''''''''''''''''
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub ImprimePlanillaBilletaje(psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Billetaje", xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = (UBound(laSdoEnc, 1) * 20) + 18   '198
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''      lsTitRp1 = "C A J A    M O N E D A    N A C I O N A L"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "C A J A    M O N E D A    E X T R A N J E R A"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!Codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 1) = "CAJA GENERAL"
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 2) = "TOTAL"
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoBil, 2) To UBound(laSdoBil, 2)
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoBil(UBound(laSdoBil, 1), i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = 0
''''''''''''''''''''''''        For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
''''''''''''''''''''''''
''''''''''''''''''''''''            xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoBil(K, i))
''''''''''''''''''''''''
''''''''''''''''''''''''            nTotAge = nTotAge + nVal(laSdoBil(K, i))
''''''''''''''''''''''''        Next
''''''''''''''''''''''''
''''''''''''''''''''''''        If psMoneda = "2" Then
''''''''''''''''''''''''            laProvis(6, i) = nTotAge
''''''''''''''''''''''''            laTotPro(6) = laTotPro(6) + nTotAge
''''''''''''''''''''''''        End If
''''''''''''''''''''''''
''''''''''''''''''''''''        'Totales por Dia
''''''''''''''''''''''''
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, 1) = "Totales"
''''''''''''''''''''''''
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''    For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
''''''''''''''''''''''''        'Totales por Agencia
''''''''''''''''''''''''
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotBil(K)
''''''''''''''''''''''''        nTotAge = nTotAge + laTotBil(K)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells.Select
''''''''''''''''''''''''    xlHoja1.Cells.Font.Name = "Arial"
''''''''''''''''''''''''    xlHoja1.Cells.Font.Size = 9
''''''''''''''''''''''''    xlHoja1.Cells.EntireColumn.AutoFit
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    If Day(ldFecFin + 1) = 1 Then  'Es ultimo dia del mes
''''''''''''''''''''''''       oEst.EliminaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda
''''''''''''''''''''''''       oEst.InsertaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda, Round(nTotAge / Day(ldFecFin), 2)
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    Set oEst = Nothing
''''''''''''''''''''''''
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub ImprimePlanillaProvision(psMoneda As String, lsFecha As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
''''''''''''''''''''''''    Dim i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''
''''''''''''''''''''''''    'Variables Calculo de Encaje
''''''''''''''''''''''''    Dim R             As ADODB.Recordset
''''''''''''''''''''''''    Dim MatTemp()     As String
''''''''''''''''''''''''    Dim MatTemp2()    As String
''''''''''''''''''''''''    Dim MatTemp3()    As String
''''''''''''''''''''''''    Dim Total1        As Double
''''''''''''''''''''''''    Dim Total2        As Double
''''''''''''''''''''''''    Dim lnTotalEncaje As Double
''''''''''''''''''''''''    Dim pdFecIni As Date
''''''''''''''''''''''''    Dim oPara As New NEncajeBCR
''''''''''''''''''''''''    Dim lnLinPag As Long
''''''''''''''''''''''''
''''''''''''''''''''''''    pdFecIni = Format("01" & Mid(lsFecha, 3), "MM/dd/YYYY")
''''''''''''''''''''''''
''''''''''''''''''''''''
'''''''''''''''''''''''''    'Variables Calculo Encaje M.E.
'''''''''''''''''''''''''    Dim lnSEncProm As Double, lnSEncMarg As Double
'''''''''''''''''''''''''    Dim lnTasaBase As Double, lnTasaEncMarg As Double
'''''''''''''''''''''''''    Dim lnEncMarg As Double, lnEncCalc As Double
'''''''''''''''''''''''''    Dim lnTObligEnc As Double, lnEncExig As Double
'''''''''''''''''''''''''    Dim lnTasaEncLegMN As Double
'''''''''''''''''''''''''    Dim lnEncLeg As Double
'''''''''''''''''''''''''    Dim lnLinPag As Integer
'''''''''''''''''''''''''    Dim lnEncMarColoc As Double
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Planilla", xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = 80
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oBala As NBalanceCont
''''''''''''''''''''''''    Set oBala = New NBalanceCont
''''''''''''''''''''''''
''''''''''''''''''''''''    'lnTasaEncLegMN = oBala.GetUtilidadAcumulada("P", -1, Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA NACIONAL"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA EXTRANJERA"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 8) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 1) = "FECHA"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 1)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 2) = "DEPOSITOS Y OBLIGACIONES  SUJETOS DE ENCAJE": xlHoja1.Cells(7, 2) = "AHORRO": xlHoja1.Cells(7, 3) = "PLAZO FIJO": xlHoja1.Cells(7, 4) = "OBLIG. INMED.": xlHoja1.Cells(7, 5) = "TOTAL"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 6) = "EFECTIVO Y DEPOSITOS DE ENCAJE":  xlHoja1.Cells(7, 6) = "CAJA": xlHoja1.Cells(7, 7) = "DEPOSITOS BCR":     xlHoja1.Cells(7, 8) = "TOTAL"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 9
''''''''''''''''''''''''    For i = LBound(laProvis, 2) To UBound(laProvis, 2) 'Recorro el Numero de Dias
''''''''''''''''''''''''
''''''''''''''''''''''''      'Dia
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag - 1, 1) = "'" & Format(laProvis(1, i), gsFormatoFechaView)
''''''''''''''''''''''''      For K = 2 To 8
''''''''''''''''''''''''          '2 = Ahorro; 3 = PF; 4 = ObligInm; 5 = Total; 6=Caja; 7= Depositos BCR; 8 =Total
''''''''''''''''''''''''          xlHoja1.Cells(i + lnLinPag - 1, K) = nVal(laProvis(K, i))
''''''''''''''''''''''''      Next
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(i + lnLinPag - 2, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    'Total de Todos los Dias
''''''''''''''''''''''''    For K = 2 To 8
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag - 1, K) = laTotPro(K)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag - 1, 1), xlHoja1.Cells(i + lnLinPag - 1, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(i + lnLinPag - 1, 8)).NumberFormat = "#,##0.00"
''''''''''''''''''''''''
''''''''''''''''''''''''    'Imprimimos Cómputo del Encaje
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 2) = "DIAS"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 3) = "ENCAJE CMAC"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 4) = "ENCAJE LEGAL"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 5) = "EXCEDENTE/DEFICIT"
''''''''''''''''''''''''
''''''''''''''''''''''''    'Calculo del Encaje
''''''''''''''''''''''''    '''''''''''''''''''
'''''''''''''''''''''''''''''    If psMoneda = "1" Then
'''''''''''''''''''''''''''''      lnEncLeg = Round((laTotPro(5)) * lnTasaEncLegMN / 100, 2) & "   "
'''''''''''''''''''''''''''''    Else
'''''''''''''''''''''''''''''      lnTObligEnc = oBala.GetUtilidadAcumulada("T", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''      If Format(lsFecha, "YYYYMMdd") >= "20040401" And Format(lsFecha, "YYYYMMdd") <= "20041031" Then
'''''''''''''''''''''''''''''        lnSEncProm = Round(lnTObligEnc / 29, 2)
'''''''''''''''''''''''''''''      ElseIf Format(lsFecha, "YYYYMMdd") >= "20041101" Then
'''''''''''''''''''''''''''''        lnSEncProm = Round(lnTObligEnc / 30, 2)
'''''''''''''''''''''''''''''      Else
'''''''''''''''''''''''''''''        lnSEncProm = Round(lnTObligEnc / 31, 2)
'''''''''''''''''''''''''''''      End If
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''      lnSEncMarg = Round(((laTotPro(5) - lnTotalColocFSD) / nDias) - lnSEncProm, 2)
'''''''''''''''''''''''''''''      lnEncExig = oBala.GetUtilidadAcumulada("E", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''''''''''''''''''''''''''      lnTasaBase = Round(lnEncExig / lnTObligEnc, 6)
'''''''''''''''''''''''''''''      If CDate(lsFecha) >= CDate("01/09/2000") And CDate(lsFecha) <= CDate("30/11/2000") Then
'''''''''''''''''''''''''''''            lnTasaBase = lnTasaBase - 0.03
'''''''''''''''''''''''''''''      End If
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''      lnTasaEncMarg = oBala.GetUtilidadAcumulada("P", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True) / 100
'''''''''''''''''''''''''''''      'lnTasaEncMarg = 0.2
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''      If lnSEncMarg > 0 Then
'''''''''''''''''''''''''''''         lnEncMarg = Round(lnSEncMarg * lnTasaEncMarg, 2)
'''''''''''''''''''''''''''''         lnEncMarColoc = Round(((lnTotalColocFSD) * lnTasaEncMarg), 2)
'''''''''''''''''''''''''''''         lnEncCalc = Round(lnSEncProm * lnTasaBase, 2)
'''''''''''''''''''''''''''''      Else
'''''''''''''''''''''''''''''         lnEncMarg = 0
'''''''''''''''''''''''''''''         lnEncMarColoc = Round(((lnTotalColocFSD) * lnTasaEncMarg), 2)
'''''''''''''''''''''''''''''         lnEncCalc = Round(Round(((laTotPro(5) - lnTotalColocFSD) / nDias), 2) * lnTasaBase, 2)
'''''''''''''''''''''''''''''      End If
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''      lnEncLeg = Round((lnEncCalc + lnEncMarg) * nDias, 2) + lnEncMarColoc
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    End If
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    'Dias
'''''''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 2) = nDias
'''''''''''''''''''''''''''''    cCodigo1 = "P_DIAS"
'''''''''''''''''''''''''''''    cCodigo2 = ""
'''''''''''''''''''''''''''''    cCodigo3 = ""
'''''''''''''''''''''''''''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Trim(Str(nDias))
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    'Encaje CMAC
'''''''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 3) = laTotPro(8)
'''''''''''''''''''''''''''''    cCodigo1 = "P_ENCAJECMAC"
'''''''''''''''''''''''''''''    cCodigo2 = ""
'''''''''''''''''''''''''''''    cCodigo3 = ""
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(laTotPro(8), "0.00")
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    'Encaje Legal
'''''''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
'''''''''''''''''''''''''''''    cCodigo1 = "P_ENCAJELEGAL"
'''''''''''''''''''''''''''''    cCodigo2 = ""
'''''''''''''''''''''''''''''    cCodigo3 = ""
'''''''''''''''''''''''''''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(Round(lnEncLeg, 2), "0.00")
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''    'Excedente/Deficit ==> Es el encaje_Acumula que se grababa antes
'''''''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
'''''''''''''''''''''''''''''    cCodigo1 = "P_EXCEDENTEDEFICIT"
'''''''''''''''''''''''''''''    cCodigo2 = ""
'''''''''''''''''''''''''''''    cCodigo3 = ""
'''''''''''''''''''''''''''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(Round(laTotPro(8) - lnEncLeg, 2), "0.00")
''''''''''''''''''''''''
''''''''''''''''''''''''    If psMoneda = "2" Then
''''''''''''''''''''''''        Set R = oPara.GetParametroEncaje(pdFecIni)
''''''''''''''''''''''''        ReDim MatTemp(UBound(laProvis, 2))
''''''''''''''''''''''''        ReDim MatTemp2(UBound(laProvis, 2))
''''''''''''''''''''''''        Total1 = 0
''''''''''''''''''''''''        Total2 = 0
''''''''''''''''''''''''        R.Find "nCodigo=2"
''''''''''''''''''''''''
''''''''''''''''''''''''        'Hallo Exceso diario EncajeActual - Base
''''''''''''''''''''''''        For i = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''            MatTemp(i) = nVal(laProvis(5, i)) - R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        Total1 = R!nValor
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Find "nCodigo=3"
''''''''''''''''''''''''        'Hallo El Basico : Base * TasaBase (21.6478)%
''''''''''''''''''''''''        For i = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''            MatTemp2(i) = Total1 * R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''        Total1 = 0#
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Find "nCodigo=4"
''''''''''''''''''''''''
''''''''''''''''''''''''        'Encaje Marginal: Marginal por el Exeso * 20%
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            MatTemp(i) = CDbl(MatTemp(i)) * R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        'Total Encaje : En Basico mas Marginal
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            MatTemp2(i) = CDbl(MatTemp(i)) + CDbl(MatTemp2(i))
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            Total1 = Total1 + CDbl(MatTemp2(i))
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''        Total1 = CDbl(Format(Total1, "#0.00"))
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Close
''''''''''''''''''''''''        Set R = Nothing
''''''''''''''''''''''''    Else
''''''''''''''''''''''''        Set R = oPara.GetParametroEncaje(pdFecIni)
''''''''''''''''''''''''        R.Find "nCodigo=1"
''''''''''''''''''''''''        Total2 = R!nValor
''''''''''''''''''''''''        R.Close
''''''''''''''''''''''''        Set R = Nothing
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''       lnTotalEncaje = CCur(laTotPro(8)) * Total2
''''''''''''''''''''''''   Else
''''''''''''''''''''''''       lnTotalEncaje = Total1
''''''''''''''''''''''''   End If
''''''''''''''''''''''''
''''''''''''''''''''''''    'Dias
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 2) = nDias
''''''''''''''''''''''''    'Encaje CMAC
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 3) = laTotPro(8)
''''''''''''''''''''''''    'Encaje Legal
''''''''''''''''''''''''    'xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnTotalEncaje, 2)
''''''''''''''''''''''''    'Excedente
''''''''''''''''''''''''    'xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnTotalEncaje, 2)
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 2, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).NumberFormat = "#,##0.00"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 3), xlHoja1.Cells(i + lnLinPag + 2, 4)).ColumnWidth = 14
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 5), xlHoja1.Cells(i + lnLinPag + 2, 5)).ColumnWidth = 18
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells.Select
''''''''''''''''''''''''    xlHoja1.Cells.Font.Name = "Arial"
''''''''''''''''''''''''    xlHoja1.Cells.Font.Size = 9
''''''''''''''''''''''''    xlHoja1.Cells.EntireColumn.AutoFit
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    'oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
''''''''''''''''''''''''    'oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnEncLeg, 2)
''''''''''''''''''''''''    oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
''''''''''''''''''''''''    oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnTotalEncaje, 2)
''''''''''''''''''''''''
''''''''''''''''''''''''    Set oEst = Nothing
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''    Set oBala = Nothing
''''''''''''''''''''''''
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
'''''''''''''''''''''''''Fin Agregado

''''''''''''''''''''''''Private Sub CalculaBilleteAge(psMoneda As String, psAgeCod As String, psCodCta As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
''''''''''''''''''''''''   Dim lsQryEnc As String
''''''''''''''''''''''''   Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''   Dim lnDiaMes As Integer
''''''''''''''''''''''''   Dim sCtaCod As String
''''''''''''''''''''''''   Dim lnSaldo As Currency
''''''''''''''''''''''''   sCtaCod = psCodCta
''''''''''''''''''''''''   Dim oConect As New DConecta
''''''''''''''''''''''''   oConect.AbreConexion
''''''''''''''''''''''''
''''''''''''''''''''''''   lsQryEnc = "SELECT Fec.dFecha, nMonto = ISNULL(( SELECT SUM(nMonto) "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "                  FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "                  WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo me1 ON me1.nMovNro = m1.nMovNro WHERE LEFT(me1.cEfectivoCod,1) = " & psMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  '" & psAgeCod & "' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) and m1.nMovEstado = " & gMovEstContabNoContable & " "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "                                and m1.nMovFlag = " & gMovFlagVigente & " ) and "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "       LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & "       SubString(m.cMovNro,18,2) = '" & psAgeCod & "' ),0)  FROM FechaTmp('" & Format(pdFecIni, gsFormatoFecha) & "')  Fec  "
''''''''''''''''''''''''
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & " WHERE CONVERT(VARCHAR(8), Fec.dFecha, 112) >='" & Format(pdFecIni, "YYYYMMdd") & "' "
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & " And CONVERT(VARCHAR(8), Fec.dFecha, 112) <='" & Format(pdFecFin, "YYYYmmdd") & "' "
''''''''''''''''''''''''
''''''''''''''''''''''''   lsQryEnc = lsQryEnc & " ORDER BY Fec.dFecha  "
''''''''''''''''''''''''   Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
''''''''''''''''''''''''    lnSaldo = 0
''''''''''''''''''''''''    lnDiaMes = 0
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''       lnDiaMes = lnDiaMes + 1
''''''''''''''''''''''''       lnSaldo = rsQryEnc!nMonto
''''''''''''''''''''''''       laSdoBil(pnContAge, lnDiaMes) = lnSaldo
''''''''''''''''''''''''       laTotBil(pnContAge) = laTotBil(pnContAge) + Val(laSdoBil(pnContAge, lnDiaMes))
''''''''''''''''''''''''       rsQryEnc.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''   oConect.CierraConexion
''''''''''''''''''''''''   Set oConect = Nothing
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaBilleteAgePromedio(psMoneda As String, psAgeCod As String, psCodCta As String, pdFecIni As Date, pdFecFin As Date)
''''''''''''''''''''''''   Dim lsQryEnc As String
''''''''''''''''''''''''   Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''   Dim lnDiaMes As Integer
''''''''''''''''''''''''   Dim lnAgeCod As Integer
''''''''''''''''''''''''   Dim sCtaCod As String
''''''''''''''''''''''''   Dim lnSaldo As Currency
''''''''''''''''''''''''   Dim lnDiaIni As Integer
''''''''''''''''''''''''   sCtaCod = psCodCta
''''''''''''''''''''''''   lnAgeCod = Val(psAgeCod)
''''''''''''''''''''''''   Dim oConect As New DConecta
''''''''''''''''''''''''   oConect.AbreConexion
'''''''''''''''''''''''''Centralizado
''''''''''''''''''''''''   lsQryEnc = "SELECT Fec.dFecha, SubString(m.cMovNro,18,2) cAgeCod, SUM(nMonto) nMonto " _
''''''''''''''''''''''''         & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
''''''''''''''''''''''''         & "       FechaTmp('" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "')  Fec " _
''''''''''''''''''''''''         & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Fec.dFecha BETWEEN '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "' and '" & Format(ldFecIni - 1, gsFormatoFecha) & "' and " _
''''''''''''''''''''''''         & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 WHERE m1.cOpeCod = m.cOpeCod and SubString(m1.cMovNro,18,2) =  '" & psAgeCod & "' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
''''''''''''''''''''''''         & "       LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
''''''''''''''''''''''''         & "       SubString(m.cMovNro,18,2) = '" & psAgeCod & "' " _
''''''''''''''''''''''''         & "GROUP BY Fec.dFecha, SubString(m.cMovNro,18,2) " _
''''''''''''''''''''''''         & "ORDER BY Fec.dFecha"
''''''''''''''''''''''''
''''''''''''''''''''''''   Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
''''''''''''''''''''''''    lnSaldo = 0
''''''''''''''''''''''''    lnDiaMes = 0
''''''''''''''''''''''''    lnDiaIni = 1
''''''''''''''''''''''''    If Not rsQryEnc.EOF Then
''''''''''''''''''''''''        lnDiaIni = Day(rsQryEnc!dFecha)
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''       lnDiaMes = lnDiaMes + 1
''''''''''''''''''''''''       lnSaldo = lnSaldo + rsQryEnc!nMonto
'''''''''''''''''''''''''       lnSaldo = lnSaldo + nval(IIf(psMoneda = "1", rsQryEnc!nMonto, rsQryEnc!nCtaSaldoImporteME))
''''''''''''''''''''''''       rsQryEnc.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    If lnDiaMes > 0 Then
''''''''''''''''''''''''        lnSaldo = Round(lnSaldo / lnDiaMes, 2)
''''''''''''''''''''''''        For lnDiaMes = lnDiaIni To UBound(laSdoBil, 2)
''''''''''''''''''''''''            'laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
''''''''''''''''''''''''            'laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
''''''''''''''''''''''''            'laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
''''''''''''''''''''''''            laProvis(6, lnDiaMes) = laProvis(6, lnDiaMes) + lnSaldo
''''''''''''''''''''''''            laTotPro(6) = laTotPro(6) + lnSaldo
''''''''''''''''''''''''        Next
''''''''''''''''''''''''    End If
''''''''''''''''''''''''   oConect.CierraConexion
''''''''''''''''''''''''   Set oConect = Nothing
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaBilleteTotCajaPromedio(psMoneda As String, psCtaCod As String, pdFecIni As Date)
''''''''''''''''''''''''Dim lsQryEnc As String
''''''''''''''''''''''''Dim rsQryEnc As ADODB.Recordset
''''''''''''''''''''''''Dim lnDiaMes As Integer
''''''''''''''''''''''''Dim lnSaldo As Currency
''''''''''''''''''''''''Dim lnTotal As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    lnSaldo = 0
''''''''''''''''''''''''    lnDiaMes = 0
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    lsQryEnc = oEst.GetImporteEstadAnexos(pdFecIni - 1, "CAJA_PROMEDIO", psMoneda)
''''''''''''''''''''''''    If lsQryEnc = "0" Then
''''''''''''''''''''''''
''''''''''''''''''''''''    'Calcula el Promedio en funcion al Total Caja / Nro Dias
''''''''''''''''''''''''      lsQryEnc = "SELECT ROUND(SUM(nMontoMN)/Count(*),2) nMontoMN, ROUND(SUM(nMontoME)/Count(*),2) nMontoME " _
''''''''''''''''''''''''            & "FROM (SELECT dFecha, SUM(nMontoMN) nMontoMN, SUM(nMontoME) nMontoME " _
''''''''''''''''''''''''            & "      FROM ( SELECT Fec.dFecha, SUM(nMonto) nMontoMN, SUM(nMonto) nMontoME " _
''''''''''''''''''''''''            & "        FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
''''''''''''''''''''''''            & "            FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "')  Fec " _
''''''''''''''''''''''''            & "        WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Datediff(m,Fec.dFecha, '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "') = 0 and " _
''''''''''''''''''''''''            & "            LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo me1 ON me1.nMovNro = m1.nMovNro WHERE m1.nMovEstado = " & gMovEstContabNoContable & " and m1.nMovFlag = " & gMovFlagVigente & " and LEFT(me1.cEfectivoCod,1) = " & psMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  SubString(m.cMovNro,18,2) and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
''''''''''''''''''''''''            & "            LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
''''''''''''''''''''''''            & "        GROUP BY Fec.dFecha UNION ALL " _
''''''''''''''''''''''''            & "        SELECT Fec.dFecha, SUM(nCtaSaldoImporte) nMontoMN, SUM(nCtaSaldoImporteME) nMontoME  " _
''''''''''''''''''''''''            & "         FROM CTASALDO CS1, FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') Fec " _
''''''''''''''''''''''''            & "         WHERE cs1.cCtaContCod = '" & psCtaCod & "' and datediff(m,Fec.dFecha,'" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') = 0 and " _
''''''''''''''''''''''''            & "             cs1.dCtaSaldoFecha = ( " _
''''''''''''''''''''''''            & "                 SELECT  MAX(dCtaSaldoFecha) " _
''''''''''''''''''''''''            & "                 FROM    CtaSaldo CS " _
''''''''''''''''''''''''            & "                 WHERE   cs1.cCtaContCod = CS.cCtaContCod AND CS.dCtaSaldoFecha <= Fec.dFecha ) " _
''''''''''''''''''''''''            & "         GROUP BY Fec.dFecha " _
''''''''''''''''''''''''            & "      ) Caja GROUP BY dFecha ) C "
''''''''''''''''''''''''
''''''''''''''''''''''''       Dim oConect As New DConecta
''''''''''''''''''''''''       oConect.AbreConexion
''''''''''''''''''''''''       Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
''''''''''''''''''''''''        If Not rsQryEnc.EOF Then
''''''''''''''''''''''''           lnDiaMes = lnDiaMes + 1
''''''''''''''''''''''''           lnSaldo = IIf(IsNull(rsQryEnc!NMONTOMN), 0, rsQryEnc!NMONTOMN)
''''''''''''''''''''''''           oEst.InsertaEstadAnexos pdFecIni - 1, "CAJA_PROMEDIO", psMoneda, Str(lnSaldo)
''''''''''''''''''''''''           rsQryEnc.MoveNext
''''''''''''''''''''''''        End If
''''''''''''''''''''''''    Else
''''''''''''''''''''''''        lnDiaMes = lnDiaMes + 1
''''''''''''''''''''''''        lnSaldo = nVal(lsQryEnc)
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    If lnDiaMes > 0 Then
''''''''''''''''''''''''        For lnDiaMes = 1 To UBound(laSdoBil, 2)
''''''''''''''''''''''''            laProvis(6, lnDiaMes) = laProvis(6, lnDiaMes) + lnSaldo
''''''''''''''''''''''''            laTotPro(6) = laTotPro(6) + lnSaldo
''''''''''''''''''''''''        Next
''''''''''''''''''''''''    End If
''''''''''''''''''''''''   oConect.CierraConexion
''''''''''''''''''''''''   Set oConect = Nothing
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaBilleteCaja(psMoneda As String, psCodCtaCaja As String, pdFecIni As Date, pdFecFin As Date)
''''''''''''''''''''''''    Dim lsQryEnc As String
''''''''''''''''''''''''    Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''    Dim lnDiaMes As Integer
''''''''''''''''''''''''    Dim lnAgeCod As Integer
''''''''''''''''''''''''    Dim lnSaldo  As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oSdo As New NCtasaldo
''''''''''''''''''''''''    RaiseEvent BarraShow(1)
''''''''''''''''''''''''    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
''''''''''''''''''''''''    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaCaja, pdFecIni, pdFecFin)
''''''''''''''''''''''''    lnSaldo = 0
''''''''''''''''''''''''    lnDiaMes = 0
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''      lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
''''''''''''''''''''''''      lnSaldo = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtaSaldoImporteME))
''''''''''''''''''''''''      laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
''''''''''''''''''''''''      laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
''''''''''''''''''''''''      laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
''''''''''''''''''''''''      rsQryEnc.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
''''''''''''''''''''''''    RaiseEvent BarraClose
''''''''''''''''''''''''    Set oSdo = New NCtasaldo
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaBilleteCajaPromedio(psMoneda As String, psCodCtaCaja As String, pdFecIni As Date, pdFecFin As Date)
''''''''''''''''''''''''    Dim lsQryEnc As String
''''''''''''''''''''''''    Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''    Dim lnDiaMes As Integer
''''''''''''''''''''''''    Dim lnAgeCod As Integer
''''''''''''''''''''''''    Dim lnSaldo  As Currency
''''''''''''''''''''''''    Dim lnTotal As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oSdo As New NCtasaldo
''''''''''''''''''''''''    RaiseEvent BarraShow(1)
''''''''''''''''''''''''    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
''''''''''''''''''''''''    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaCaja, DateAdd("m", -1, ldFecIni), ldFecIni - 1)
''''''''''''''''''''''''    lnSaldo = 0
''''''''''''''''''''''''    lnDiaMes = 0
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''       lnDiaMes = lnDiaMes + 1
''''''''''''''''''''''''       lnSaldo = lnSaldo + nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtaSaldoImporteME))
''''''''''''''''''''''''       rsQryEnc.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    If lnDiaMes > 0 Then
''''''''''''''''''''''''        lnSaldo = Round(lnSaldo / lnDiaMes, 2)
''''''''''''''''''''''''        lnTotal = 0
''''''''''''''''''''''''        For lnDiaMes = 1 To UBound(laSdoBil, 2)
'''''''''''''''''''''''''            laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
'''''''''''''''''''''''''            laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
'''''''''''''''''''''''''            laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
''''''''''''''''''''''''            lnTotal = lnTotal + lnSaldo
''''''''''''''''''''''''            laProvis(6, lnDiaMes) = lnSaldo
''''''''''''''''''''''''        Next
''''''''''''''''''''''''        laTotPro(6) = laTotPro(6) + lnTotal
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
''''''''''''''''''''''''    RaiseEvent BarraClose
''''''''''''''''''''''''    Set oSdo = New NCtasaldo
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaSaldoBCR(psMoneda As String, psCodCtaBCR As String, pdFecIni As Date, pdFecFin As Date)
''''''''''''''''''''''''    Dim lsQryEnc As String
''''''''''''''''''''''''    Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''    Dim lnDiaMes As Integer
''''''''''''''''''''''''    Dim lnSaldo  As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oSdo As New NCtasaldo
''''''''''''''''''''''''    RaiseEvent BarraShow(1)
''''''''''''''''''''''''    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS BANCO BCR", "Procesando...", vbBlue)
''''''''''''''''''''''''    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaBCR & "%", pdFecIni, pdFecFin)
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''       lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
''''''''''''''''''''''''       laProvis(1, lnDiaMes) = Format(rsQryEnc!dFecha, gsFormatoFechaView)
''''''''''''''''''''''''       laProvis(7, lnDiaMes) = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtaSaldoImporteME))
''''''''''''''''''''''''       rsQryEnc.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    For lnDiaMes = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''      If lnDiaMes = 1 Then
''''''''''''''''''''''''         If Val(laProvis(7, 1)) = 0 Then
''''''''''''''''''''''''            lnSaldo = oSdo.GetCtaSaldo(psCodCtaBCR, Format(pdFecIni, gsFormatoFecha), psMoneda = "1")
''''''''''''''''''''''''            laProvis(1, lnDiaMes) = pdFecIni
''''''''''''''''''''''''            laProvis(7, lnDiaMes) = lnSaldo
''''''''''''''''''''''''         End If
''''''''''''''''''''''''      Else
''''''''''''''''''''''''         If Val(laProvis(7, lnDiaMes)) = 0 Then
''''''''''''''''''''''''            laProvis(1, lnDiaMes) = pdFecIni + lnDiaMes - 1
''''''''''''''''''''''''            laProvis(7, lnDiaMes) = laProvis(7, lnDiaMes - 1)
''''''''''''''''''''''''         End If
''''''''''''''''''''''''      End If
''''''''''''''''''''''''      laTotPro(7) = laTotPro(7) + Val(laProvis(7, lnDiaMes))
''''''''''''''''''''''''    Next
''''''''''''''''''''''''    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS BANCO BCR", "Procesando...", vbBlue)
''''''''''''''''''''''''    Set oSdo = New NCtasaldo
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaSdoEncObligaciones(psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim dTempo As Date
''''''''''''''''''''''''    Dim lnDiaMes As Integer
''''''''''''''''''''''''    Dim Total As Double
''''''''''''''''''''''''
''''''''''''''''''''''''    RaiseEvent BarraShow(1)
''''''''''''''''''''''''    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS OBLIGACIONES INMEDIATAS", "Procesando..", vbBlue)
''''''''''''''''''''''''
''''''''''''''''''''''''    dTempo = pdFecIni
''''''''''''''''''''''''
''''''''''''''''''''''''    Do While Format(pdFecFin, "YYYYmmdd") >= Format(dTempo, "YYYYmmdd")
''''''''''''''''''''''''
''''''''''''''''''''''''        lnDiaMes = dTempo - pdFecIni + 1
''''''''''''''''''''''''
''''''''''''''''''''''''        Total = SaldoObligInmediatas(1, "76" & psMoneda & "201", dTempo, pdFecFin, psAgeCod)
''''''''''''''''''''''''        laSdoEncInt(pnContAge, 1, lnDiaMes) = Format(dTempo, gsFormatoFechaView)
''''''''''''''''''''''''        laSdoEncInt(pnContAge, 2, lnDiaMes) = Total
''''''''''''''''''''''''        laTotEncInt(pnContAge, 1) = Val(laTotEncInt(pnContAge, 1)) + laSdoEncInt(pnContAge, 2, lnDiaMes)
''''''''''''''''''''''''
''''''''''''''''''''''''        dTempo = dTempo + 1
''''''''''''''''''''''''
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''
''''''''''''''''''''''''    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS OBLIGACIONES INMEDIATAS", "Procesando...", vbBlue)
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Public Function SaldoObligInmediatas(lnNroCol As Long, lsOpeCod As String, ByVal ldFecha As Date, ByVal pdFecFin As Date, ByVal cCodAge As String) As Currency
''''''''''''''''''''''''Dim rsC As New ADODB.Recordset
''''''''''''''''''''''''Dim oRep As New DRepCtaColumna
''''''''''''''''''''''''
''''''''''''''''''''''''Dim oTC As New nTipoCambio
''''''''''''''''''''''''Dim pnTCF As Double
''''''''''''''''''''''''Dim pnTCFF As Double
''''''''''''''''''''''''
''''''''''''''''''''''''pnTCF = Format(oTC.EmiteTipoCambio(pdFecFin, TCFijoMes), "#,##0.00###")
''''''''''''''''''''''''pnTCFF = Format(oTC.EmiteTipoCambio(pdFecFin + 1, TCFijoMes), "#,##0.00###")
''''''''''''''''''''''''
''''''''''''''''''''''''Set rsC = oRep.GetRepColumnaCtaSaldo(lsOpeCod, lnNroCol, ldFecha, gbBitCentral, cCodAge)
''''''''''''''''''''''''SaldoObligInmediatas = 0
''''''''''''''''''''''''If Not RSVacio(rsC) Then
''''''''''''''''''''''''    If Mid(lsOpeCod, 3, 1) = "1" Then
''''''''''''''''''''''''        SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, rsC!Total)
''''''''''''''''''''''''    Else
''''''''''''''''''''''''        If pdFecFin = ldFecha Then
''''''''''''''''''''''''            If pnTCFF = 0 Then
''''''''''''''''''''''''                SaldoObligInmediatas = 0
''''''''''''''''''''''''            Else
''''''''''''''''''''''''                SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCFF, 2))
''''''''''''''''''''''''            End If
''''''''''''''''''''''''        Else
''''''''''''''''''''''''            If pnTCF = 0 Then
''''''''''''''''''''''''                SaldoObligInmediatas = 0
''''''''''''''''''''''''            Else
''''''''''''''''''''''''                SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCF, 2))
''''''''''''''''''''''''            End If
''''''''''''''''''''''''        End If
''''''''''''''''''''''''    End If
''''''''''''''''''''''''End If
''''''''''''''''''''''''Set oRep = Nothing
''''''''''''''''''''''''RSClose rsC
''''''''''''''''''''''''End Function
''''''''''''''''''''''''
'''''''''''''''''''''''''Private Sub CalculaSdoEncInteres(psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
'''''''''''''''''''''''''    Dim lsQryEnc As String
'''''''''''''''''''''''''    Dim rsQryEnc As New ADODB.Recordset
'''''''''''''''''''''''''    Dim lnDiaMes As Integer
'''''''''''''''''''''''''    Dim ldFecha     As Date
'''''''''''''''''''''''''    Dim lnSaldo     As Currency
'''''''''''''''''''''''''    Dim i           As Integer
'''''''''''''''''''''''''    Dim lsCtaCodInt As String
'''''''''''''''''''''''''    Dim rsCta As ADODB.Recordset
'''''''''''''''''''''''''
'''''''''''''''''''''''''    ldFecha = ldFecIni
'''''''''''''''''''''''''    Dim oDOpe As DOperacion
'''''''''''''''''''''''''    Dim oSdo  As NCtasaldo
'''''''''''''''''''''''''    RaiseEvent BarraShow(1)
'''''''''''''''''''''''''    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS ENCAJE INTERESES", "Procesando..", vbBlue)
'''''''''''''''''''''''''    Set oDOpe = New DOperacion
'''''''''''''''''''''''''    Set oSdo = New NCtasaldo
'''''''''''''''''''''''''    'lsCtaCodInt = oDOpe.EmiteOpeCta(lsOpeCod, "D", "3")
'''''''''''''''''''''''''    Set rsCta = oDOpe.CargaOpeCtaUltimoNivel(lsOpeCod, "D", 3)
'''''''''''''''''''''''''    Do While Not rsCta.EOF
'''''''''''''''''''''''''        If Right(rsCta!cCtaContCod, 2) = psAgeCod Then
'''''''''''''''''''''''''            lsCtaCodInt = lsCtaCodInt & "'" & rsCta!cCtaContCod & "',"
'''''''''''''''''''''''''        End If
'''''''''''''''''''''''''        rsCta.MoveNext
'''''''''''''''''''''''''    Loop
'''''''''''''''''''''''''    If Not lsCtaCodInt = "" Then
'''''''''''''''''''''''''        lsCtaCodInt = Left(lsCtaCodInt, Len(lsCtaCodInt) - 1)
'''''''''''''''''''''''''    End If
'''''''''''''''''''''''''    RSClose rsCta
'''''''''''''''''''''''''
'''''''''''''''''''''''''    Set rsQryEnc = oSdo.GetCtaSaldoRango(lsCtaCodInt, pdFecIni, pdFecFin)
'''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
'''''''''''''''''''''''''       lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
'''''''''''''''''''''''''       laSdoEncInt(pnContAge, 1, lnDiaMes) = Format(rsQryEnc!dFecha, gsFormatoFechaView)
'''''''''''''''''''''''''       laSdoEncInt(pnContAge, 2, lnDiaMes) = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtaSaldoImporteME))
'''''''''''''''''''''''''       laTotEncInt(pnContAge, 1) = Val(laTotEncInt(pnContAge, 1)) + laSdoEncInt(pnContAge, 2, lnDiaMes)
'''''''''''''''''''''''''       rsQryEnc.MoveNext
'''''''''''''''''''''''''    Loop
'''''''''''''''''''''''''    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS ENCAJE INTERESES", "Procesando...", vbBlue)
'''''''''''''''''''''''''    Set oSdo = Nothing
'''''''''''''''''''''''''    Set oDOpe = Nothing
'''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaSdoEnc(pnBitCentral_ As Boolean, psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
''''''''''''''''''''''''    Dim lsQryEnc As String
''''''''''''''''''''''''    Dim rsQryEnc As New ADODB.Recordset
''''''''''''''''''''''''    Dim lnDiaMes As Integer
''''''''''''''''''''''''    Dim pdFechaControl As Date
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oConect As New DConecta
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    oConect.AbreConexion
''''''''''''''''''''''''    pdFechaControl = pdFecIni
''''''''''''''''''''''''    Set rsQryEnc = oEst.GetEstadisticaAhorro(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''        lnDiaMes = rsQryEnc!dEstadAC - pdFecIni + 1
''''''''''''''''''''''''        If Day(pdFechaControl) <> lnDiaMes Then
'''''''''''''''''''''''''          MsgBox "No se encontraron datos Estadisticos AC del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
''''''''''''''''''''''''        Else
''''''''''''''''''''''''            If pnBitCentral_ = True Then
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) '- nVal(rsQryEnc!nSaldCRAC)
''''''''''''''''''''''''                Else
''''''''''''''''''''''''                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) + nVal(rsQryEnc!nSaldoCMAC) + nVal(rsQryEnc!nSaldoCRAC) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            Else
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoAC) + IIf(pnContAge = 1, nPPD, 0) - nVal(rsQryEnc!nSaldoCMAC) - nVal(rsQryEnc!nSaldCRAC)
''''''''''''''''''''''''                Else
''''''''''''''''''''''''                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoAC) + IIf(pnContAge = 1, nPPD, 0) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            End If
''''''''''''''''''''''''
''''''''''''''''''''''''            laTotEnc(pnContAge, 2) = Val(laTotEnc(pnContAge, 2)) + Val(laSdoEnc(pnContAge, 2, lnDiaMes))
''''''''''''''''''''''''            rsQryEnc.MoveNext
''''''''''''''''''''''''        End If
''''''''''''''''''''''''        pdFechaControl = pdFechaControl + 1
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rsQryEnc
''''''''''''''''''''''''    pdFechaControl = pdFecIni
''''''''''''''''''''''''    Set rsQryEnc = oEst.GetEstadisticaPlazoFijo(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''        lnDiaMes = rsQryEnc!dEstadPF - pdFecIni + 1
''''''''''''''''''''''''        If Day(pdFechaControl) <> lnDiaMes Then
'''''''''''''''''''''''''          MsgBox "No se encontraron datos Estadisticos PF del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
''''''''''''''''''''''''        Else
''''''''''''''''''''''''            If pnBitCentral_ = True Then
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac)
''''''''''''''''''''''''                Else
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) + nVal(rsQryEnc!nSaldoCMAC) + nVal(rsQryEnc!nSaldoCRAC) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            Else
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoPF) - nVal(rsQryEnc!nSaldoCMAC) - nVal(rsQryEnc!nSaldCRAC)
''''''''''''''''''''''''                Else
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoPF) + nVal(rsQryEnc!nMonChqVal)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            End If
''''''''''''''''''''''''            laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + Val(laSdoEnc(pnContAge, 3, lnDiaMes))
''''''''''''''''''''''''            rsQryEnc.MoveNext
''''''''''''''''''''''''        End If
''''''''''''''''''''''''        pdFechaControl = pdFechaControl + 1
''''''''''''''''''''''''
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rsQryEnc
''''''''''''''''''''''''
''''''''''''''''''''''''    pdFechaControl = pdFecIni
''''''''''''''''''''''''    Set rsQryEnc = oEst.GetEstadisticaCTS(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
''''''''''''''''''''''''    Do While Not rsQryEnc.EOF
''''''''''''''''''''''''        lnDiaMes = rsQryEnc!dEstadCTS - pdFecIni + 1
''''''''''''''''''''''''        If Day(pdFechaControl) <> lnDiaMes Then
'''''''''''''''''''''''''          MsgBox "No se encontraron datos Estadisticos CTS del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
''''''''''''''''''''''''        Else
''''''''''''''''''''''''            If pnBitCentral_ = True Then
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(laSdoEnc(pnContAge, 3, lnDiaMes)) + nVal(rsQryEnc!nSaldoSinCmac)
''''''''''''''''''''''''                    laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + nVal(rsQryEnc!nSaldoSinCmac)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            Else
''''''''''''''''''''''''                If lnTipo = 1 Then
''''''''''''''''''''''''                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(laSdoEnc(pnContAge, 3, lnDiaMes)) + nVal(rsQryEnc!nSaldo)
''''''''''''''''''''''''                    laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + nVal(rsQryEnc!nSaldo)
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            End If
''''''''''''''''''''''''            rsQryEnc.MoveNext
''''''''''''''''''''''''        End If
''''''''''''''''''''''''        pdFechaControl = pdFechaControl + 1
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rsQryEnc
''''''''''''''''''''''''End Sub
''''''''''''''''''''''''
''''''''''''''''''''''''Private Function PrtSdoEnc(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = ""
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = UBound(laSdoEnc, 1) * 20   '  180
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''      lsTitRp1 = "E F E C T O S   D E   E N C A J E (CAPITAL)"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    If nTipo = 2 Then
''''''''''''''''''''''''      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "SEDE", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecIni, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & Space(11)
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!codigo, "00") & "    "
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & "   T O T A L    " & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 1) = "TOTAL"
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoEnc, 3) To UBound(laSdoEnc, 3)
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoEnc(1, 1, i), gsFormatoFechaView) & " "
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''      nTotAge = 0
''''''''''''''''''''''''      For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoEnc(K, nTipo, i)), 15, 2) & " "
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoEnc(K, nTipo, i))
''''''''''''''''''''''''        nTotAge = nTotAge + nVal(laSdoEnc(K, nTipo, i))
''''''''''''''''''''''''      Next
''''''''''''''''''''''''      laProvis(1, i) = laSdoEnc(1, 1, i)
''''''''''''''''''''''''      laProvis(nTipo, i) = nTotAge
''''''''''''''''''''''''      laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & PrnVal(nTotAge, 15, 2)
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(11, " ")
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''    For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & PrnVal(laTotEnc(K, nTipo), 15, 2) & " "
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotEnc(K, nTipo)
''''''''''''''''''''''''        nTotAge = nTotAge + laTotEnc(K, nTipo)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & PrnVal(nTotAge, 15, 2) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
''''''''''''''''''''''''    PrtSdoEnc = lsRTFEnc
''''''''''''''''''''''''End Function
''''''''''''''''''''''''Private Function PrtSdoEncInt(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''    lsRTFEnc = ""
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = UBound(laSdoEncInt, 1) * 20   '  180
''''''''''''''''''''''''    If lnTipo = 1 Then
''''''''''''''''''''''''      lsTitRp1 = "E F E C T O S   D E   E N C A J E (INTERES)"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    If nTipo = 2 Then
''''''''''''''''''''''''      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecIni, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & Space(11)
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!codigo, "00") & "    "
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & "   T O T A L    " & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 1) = "TOTAL"
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoEncInt, 3) To UBound(laSdoEncInt, 3)
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoEncInt(1, 1, i), gsFormatoFechaView) & " "
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''      nTotAge = 0
''''''''''''''''''''''''      For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoEncInt(K, 2, i)), 15, 2) & "  "
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoEncInt(K, 2, i))
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = nTotAge + nVal(laSdoEncInt(K, 2, i))
''''''''''''''''''''''''      Next
''''''''''''''''''''''''      'adiciona la provision para el total en la planilla provisional de encaje
''''''''''''''''''''''''      laProvis(nTipo, i) = nTotAge
''''''''''''''''''''''''      laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2)
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(11, " ")
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''    For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & PrnVal(Str(laTotEncInt(K, 1)), 15, 2) & "  "
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotEncInt(K, 1)
''''''''''''''''''''''''
''''''''''''''''''''''''        nTotAge = nTotAge + laTotEncInt(K, 1)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
''''''''''''''''''''''''    PrtSdoEncInt = lsRTFEnc
''''''''''''''''''''''''End Function
''''''''''''''''''''''''
''''''''''''''''''''''''Private Function PrtSdoBillete(psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lnLinPag As Integer
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''    Dim nTotAge As Currency
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Billetaje", xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Orientation = xlLandscape
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = False
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesTall = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.FitToPagesWide = 1
''''''''''''''''''''''''    xlHoja1.PageSetup.CenterVertically = True
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = ""
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = (UBound(laSdoEnc, 1) * 20) + 18   '198
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''      lsTitRp1 = "C A J A    M O N E D A    N A C I O N A L"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "C A J A    M O N E D A    E X T R A N J E R A"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecFin, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & Space(11)
''''''''''''''''''''''''
'''''''''''''''''''''''''Cabecera de Agencias
''''''''''''''''''''''''    Dim oDAge As New DActualizaDatosArea
''''''''''''''''''''''''    Set rs = oDAge.GetAgencias(, False, True)
''''''''''''''''''''''''    Set oDAge = Nothing
''''''''''''''''''''''''    i = LBound(laSdoEnc, 1)
''''''''''''''''''''''''    Do While Not rs.EOF
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!codigo, "00") & "    "
''''''''''''''''''''''''        xlHoja1.Cells(6, i + 1) = "Agencia " & Format(rs!codigo, "00")
''''''''''''''''''''''''        i = i + 1
''''''''''''''''''''''''        rs.MoveNext
''''''''''''''''''''''''    Loop
''''''''''''''''''''''''    RSClose rs
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 10) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & "  CAJA GENERAL          T O T A L "
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 1) = "CAJA GENERAL"
''''''''''''''''''''''''    xlHoja1.Cells(6, i + 2) = "TOTAL"
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lnLinPag = 8
''''''''''''''''''''''''    For i = LBound(laSdoBil, 2) To UBound(laSdoBil, 2)
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoBil(UBound(laSdoBil, 1), i), gsFormatoFechaView) & " "
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, 1) = "'" & Format(laSdoBil(UBound(laSdoBil, 1), i), gsFormatoFechaView)
''''''''''''''''''''''''
''''''''''''''''''''''''      nTotAge = 0
''''''''''''''''''''''''      For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
''''''''''''''''''''''''          lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoBil(K, i)), 15, 2) & "  "
''''''''''''''''''''''''          xlHoja1.Cells(i + lnLinPag, K + 1) = nVal(laSdoBil(K, i))
''''''''''''''''''''''''
''''''''''''''''''''''''          nTotAge = nTotAge + nVal(laSdoBil(K, i))
''''''''''''''''''''''''      Next
''''''''''''''''''''''''      If psMoneda = "2" Then
''''''''''''''''''''''''         laProvis(6, i) = nTotAge
''''''''''''''''''''''''         laTotPro(6) = laTotPro(6) + nTotAge
''''''''''''''''''''''''      End If
''''''''''''''''''''''''      lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2)
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(11, " ")
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, 1) = "Totales"
''''''''''''''''''''''''
''''''''''''''''''''''''    nTotAge = 0
''''''''''''''''''''''''    For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
''''''''''''''''''''''''        lsRTFEnc = lsRTFEnc & PrnVal(Str(laTotBil(K)), 15, 2) & "  "
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag, K + 1) = laTotBil(K)
''''''''''''''''''''''''        nTotAge = nTotAge + laTotBil(K)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2) & oImpresora.gPrnSaltoLinea
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag, K + 1) = nTotAge
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(i + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag, 1), xlHoja1.Cells(i + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    If Day(ldFecFin + 1) = 1 Then  'Es ultimo dia del mes
''''''''''''''''''''''''       oEst.EliminaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda
''''''''''''''''''''''''       oEst.InsertaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda, Round(nTotAge / Day(ldFecFin), 2)
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    Set oEst = Nothing
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=") & oImpresora.gPrnCondensadaOFF
''''''''''''''''''''''''    PrtSdoBillete = lsRTFEnc
''''''''''''''''''''''''End Function
''''''''''''''''''''''''
''''''''''''''''''''''''Private Function PrtPlanillaProvision(psMoneda As String, lsFecha As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
''''''''''''''''''''''''    Dim lnCarLin As Integer, i As Integer, K As Integer
''''''''''''''''''''''''    Dim lsTitRp1 As String
''''''''''''''''''''''''    Dim lsTitRp2 As String
''''''''''''''''''''''''    Dim lsCodMon As String
''''''''''''''''''''''''
''''''''''''''''''''''''    'Variables Calculo de Encaje
''''''''''''''''''''''''    Dim R             As ADODB.Recordset
''''''''''''''''''''''''    Dim MatTemp()     As String
''''''''''''''''''''''''    Dim MatTemp2()    As String
''''''''''''''''''''''''    Dim MatTemp3()    As String
''''''''''''''''''''''''    Dim Total1        As Double
''''''''''''''''''''''''    Dim Total2        As Double
''''''''''''''''''''''''    Dim lnTotalEncaje As Double
''''''''''''''''''''''''    Dim pdFecIni As Date
''''''''''''''''''''''''    Dim oPara As New NEncajeBCR
''''''''''''''''''''''''    Dim lnLinPag As Long
''''''''''''''''''''''''
''''''''''''''''''''''''    pdFecIni = Format("01" & Mid(lsFecha, 3), "MM/dd/YYYY")
''''''''''''''''''''''''
''''''''''''''''''''''''    'Variables Calculo Encaje M.E.
'''''''''''''''''''''''''    Dim lnSEncProm As Double, lnSEncMarg As Double
'''''''''''''''''''''''''    Dim lnTasaBase As Double, lnTasaEncMarg As Double
'''''''''''''''''''''''''    Dim lnEncMarg As Double, lnEncCalc As Double
'''''''''''''''''''''''''    Dim lnTObligEnc As Double, lnEncExig As Double
''''''''''''''''''''''''
'''''''''''''''''''''''''    Dim lnTasaEncLegMN As Double
'''''''''''''''''''''''''    Dim lnEncLeg As Double
''''''''''''''''''''''''
''''''''''''''''''''''''    ExcelAddHoja "Planilla", xlLibro, xlHoja1
''''''''''''''''''''''''    xlHoja1.PageSetup.Zoom = 80
''''''''''''''''''''''''
''''''''''''''''''''''''    lsRTFEnc = ""
''''''''''''''''''''''''    lnLinPag = 65
''''''''''''''''''''''''    lnCarLin = 140
''''''''''''''''''''''''
'''''''''''''''''''''''''    Dim oBala As NBalanceCont
'''''''''''''''''''''''''    Set oBala = New NBalanceCont
'''''''''''''''''''''''''    lnTasaEncLegMN = oBala.GetUtilidadAcumulada("P", -1, Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA NACIONAL"
''''''''''''''''''''''''    Else
''''''''''''''''''''''''      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA EXTRANJERA"
''''''''''''''''''''''''    End If
''''''''''''''''''''''''    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(1, 1) = gsInstCmac
''''''''''''''''''''''''    xlHoja1.Cells(1, 8) = ldFecFin
''''''''''''''''''''''''    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
''''''''''''''''''''''''    xlHoja1.Cells(3, 2) = lsTitRp1
''''''''''''''''''''''''    xlHoja1.Cells(4, 2) = lsTitRp2
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).Font.Bold = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 1) = "FECHA"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 1)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 2) = "DEPOSITOS Y OBLIGACIONES  SUJETOS DE ENCAJE": xlHoja1.Cells(7, 2) = "AHORRO": xlHoja1.Cells(7, 3) = "PLAZO FIJO": xlHoja1.Cells(7, 4) = "INTERES": xlHoja1.Cells(7, 5) = "TOTAL"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).MergeCells = True
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Cells(6, 6) = "EFECTIVO Y DEPOSITOS DE ENCAJE":  xlHoja1.Cells(7, 6) = "CAJA": xlHoja1.Cells(7, 7) = "DEPOSITOS BCR":     xlHoja1.Cells(7, 8) = "TOTAL"
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 8)).MergeCells = True
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
''''''''''''''''''''''''
''''''''''''''''''''''''    lnLinPag = 9
''''''''''''''''''''''''
''''''''''''''''''''''''    For i = LBound(laProvis, 2) To UBound(laProvis, 2)
''''''''''''''''''''''''      xlHoja1.Cells(i + lnLinPag - 1, 1) = "'" & Format(laProvis(1, i), gsFormatoFechaView)
''''''''''''''''''''''''      For K = 2 To 8
''''''''''''''''''''''''          xlHoja1.Cells(i + lnLinPag - 1, K) = nVal(laProvis(K, i))
''''''''''''''''''''''''      Next
''''''''''''''''''''''''    Next i
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(i + lnLinPag - 2, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''
''''''''''''''''''''''''    For K = 2 To 8
''''''''''''''''''''''''        xlHoja1.Cells(i + lnLinPag - 1, K) = laTotPro(K)
''''''''''''''''''''''''    Next
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag - 1, 1), xlHoja1.Cells(i + lnLinPag - 1, 8)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(i + lnLinPag - 1, 8)).NumberFormat = "#,##0.00"
''''''''''''''''''''''''
''''''''''''''''''''''''    'Imprimimos Cómputo del Encaje
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 2) = "DIAS"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 3) = "ENCAJE CMAC"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 4) = "ENCAJE LEGAL"
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 1, 5) = "EXCEDENTE/DEFICIT"
''''''''''''''''''''''''
'''''''''''''''''''''''''    If psMoneda = "1" Then
'''''''''''''''''''''''''      lnEncLeg = Round(laTotPro(5) * lnTasaEncLegMN / 100, 2) & "   "
'''''''''''''''''''''''''    Else
'''''''''''''''''''''''''      lnTObligEnc = oBala.GetUtilidadAcumulada("T", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''''''''''''''''''''''      lnSEncProm = Round(lnTObligEnc / 31, 2)
'''''''''''''''''''''''''      lnSEncMarg = Round((laTotPro(5) / nDias) - lnSEncProm, 2)
'''''''''''''''''''''''''      lnEncExig = oBala.GetUtilidadAcumulada("E", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''''''''''''''''''''''      lnTasaBase = Round(lnEncExig / lnTObligEnc, 6)
'''''''''''''''''''''''''      If CDate(lsFecha) >= CDate("01/09/2000") And CDate(lsFecha) <= CDate("30/11/2000") Then
'''''''''''''''''''''''''            lnTasaBase = lnTasaBase - 0.03
'''''''''''''''''''''''''      End If
'''''''''''''''''''''''''      lnTasaEncMarg = 0.2
'''''''''''''''''''''''''      If lnSEncMarg > 0 Then
'''''''''''''''''''''''''         lnEncMarg = Round(lnSEncMarg * lnTasaEncMarg, 2)
'''''''''''''''''''''''''      Else
'''''''''''''''''''''''''         lnEncMarg = Round((lnSEncProm + lnSEncMarg) * lnTasaBase, 2)
'''''''''''''''''''''''''      End If
'''''''''''''''''''''''''      lnEncCalc = Round(lnSEncProm * lnTasaBase, 2)
'''''''''''''''''''''''''      lnEncLeg = Round((lnEncCalc + lnEncMarg) * nDias, 2)
'''''''''''''''''''''''''
'''''''''''''''''''''''''
'''''''''''''''''''''''''    End If
''''''''''''''''''''''''
'''''''''''''''''''''''''    For i = LBound(laProvis, 2) To UBound(laProvis, 2)
'''''''''''''''''''''''''        MsgBox nVal(laProvis(8, i))
'''''''''''''''''''''''''    Next
''''''''''''''''''''''''
''''''''''''''''''''''''    If psMoneda = "2" Then
''''''''''''''''''''''''        Set R = oPara.GetParametroEncaje(pdFecIni)
''''''''''''''''''''''''        ReDim MatTemp(UBound(laProvis, 2))
''''''''''''''''''''''''        ReDim MatTemp2(UBound(laProvis, 2))
''''''''''''''''''''''''        Total1 = 0
''''''''''''''''''''''''        Total2 = 0
''''''''''''''''''''''''        R.Find "nCodigo=2"
''''''''''''''''''''''''
''''''''''''''''''''''''        'Hallo Exceso diario EncajeActual - Base
''''''''''''''''''''''''        For i = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''            MatTemp(i) = nVal(laProvis(5, i)) - R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        Total1 = R!nValor
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Find "nCodigo=3"
''''''''''''''''''''''''        'Hallo El Basico : Base * TasaBase (21.6478)%
''''''''''''''''''''''''        For i = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''            MatTemp2(i) = Total1 * R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''        Total1 = 0#
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Find "nCodigo=4"
''''''''''''''''''''''''
''''''''''''''''''''''''        'Encaje Marginal: Marginal por el Exeso * 20%
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            MatTemp(i) = CDbl(MatTemp(i)) * R!nValor
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        'Total Encaje : En Basico mas Marginal
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            MatTemp2(i) = CDbl(MatTemp(i)) + CDbl(MatTemp2(i))
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''
''''''''''''''''''''''''        For i = 1 To UBound(MatTemp)
''''''''''''''''''''''''            Total1 = Total1 + CDbl(MatTemp2(i))
''''''''''''''''''''''''        Next i
''''''''''''''''''''''''        Total1 = CDbl(Format(Total1, "#0.00"))
''''''''''''''''''''''''
''''''''''''''''''''''''        R.Close
''''''''''''''''''''''''        Set R = Nothing
''''''''''''''''''''''''    Else
''''''''''''''''''''''''        Set R = oPara.GetParametroEncaje(pdFecIni)
''''''''''''''''''''''''        R.Find "nCodigo=1"
''''''''''''''''''''''''        Total2 = R!nValor
''''''''''''''''''''''''        R.Close
''''''''''''''''''''''''        Set R = Nothing
''''''''''''''''''''''''    End If
''''''''''''''''''''''''
''''''''''''''''''''''''    If psMoneda = "1" Then
''''''''''''''''''''''''       lnTotalEncaje = CCur(laTotPro(8)) * Total2
''''''''''''''''''''''''   Else
''''''''''''''''''''''''       lnTotalEncaje = Total1
''''''''''''''''''''''''   End If
''''''''''''''''''''''''
''''''''''''''''''''''''    'Dias
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 2) = nDias
''''''''''''''''''''''''    'Encaje CMAC
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 3) = laTotPro(8)
''''''''''''''''''''''''    'Encaje Legal
''''''''''''''''''''''''    'xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnTotalEncaje, 2)
''''''''''''''''''''''''    'Excedente
''''''''''''''''''''''''    'xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
''''''''''''''''''''''''    xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnTotalEncaje, 2)
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).BorderAround xlContinuous, xlMedium
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 2, 2), xlHoja1.Cells(i + lnLinPag + 2, 5)).NumberFormat = "#,##0.00"
''''''''''''''''''''''''
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 3), xlHoja1.Cells(i + lnLinPag + 2, 4)).ColumnWidth = 14
''''''''''''''''''''''''    xlHoja1.Range(xlHoja1.Cells(i + lnLinPag + 1, 5), xlHoja1.Cells(i + lnLinPag + 2, 5)).ColumnWidth = 18
''''''''''''''''''''''''
''''''''''''''''''''''''    Dim oEst As New NEstadisticas
''''''''''''''''''''''''    oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
''''''''''''''''''''''''    oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnTotalEncaje, 2)
''''''''''''''''''''''''    Set oEst = Nothing
''''''''''''''''''''''''
''''''''''''''''''''''''    PrtPlanillaProvision = lsRTFEnc
''''''''''''''''''''''''    'Set oBala = Nothing
''''''''''''''''''''''''End Function
''''''''''''''''''''''''
''''''''''''''''''''''''Private Sub CalculaTotProvis()
''''''''''''''''''''''''Dim K As Integer
''''''''''''''''''''''''Dim nTot1 As Currency
''''''''''''''''''''''''Dim nTot2 As Currency
''''''''''''''''''''''''laTotPro(5) = 0: laTotPro(8) = 0
''''''''''''''''''''''''For K = 1 To UBound(laProvis, 2)
''''''''''''''''''''''''   laProvis(5, K) = Val(laProvis(2, K)) + Val(laProvis(3, K) + Val(laProvis(4, K)))
''''''''''''''''''''''''   laProvis(8, K) = Val(laProvis(6, K)) + Val(laProvis(7, K))
''''''''''''''''''''''''   laTotPro(5) = laTotPro(5) + Val(laProvis(5, K))
''''''''''''''''''''''''   laTotPro(8) = laTotPro(8) + Val(laProvis(8, K))
''''''''''''''''''''''''Next
''''''''''''''''''''''''End Sub

Public Function ImprimeConsolidaSdoEnc(pnBitCentral_ As Boolean, psOpeCod As String, psFecIni As String, psFecFin As String, pnTipo As Integer, _
       xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
Dim I As Integer, J As Integer, K As Integer

Dim dFecha As Date
Dim lsCodCtaCaja As String
Dim lsCodCtaBanc As String
Dim lsCodCtaCAge As String
Dim lsMoneda     As String
Dim lnContAge      As Integer

Dim oDOpe As New DOperacion
Dim oSdo  As New NCtasaldo
Dim oDAge As New DActualizaDatosArea


Dim lnDiaMes As Integer
Dim dTempo As Date
Dim nTotal As Double

lsOpeCod = psOpeCod
lsMoneda = Mid(psOpeCod, 3, 1)
lnTipo = pnTipo
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "", "Obteniendo datos preliminares..", vbBlue)

lsCodCtaCAge = oDOpe.EmiteOpeCta(psOpeCod, "H", "0")
lsCodCtaCaja = oDOpe.EmiteOpeCta(psOpeCod, "H", "1")
lsCodCtaBanc = oDOpe.EmiteOpeCta(psOpeCod, "H", "2")

Set rs = oDAge.GetAgencias(, False, True)

lnNroAgencias = rs.RecordCount
RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "", "Obteniendo datos preliminares..", vbBlue)
RaiseEvent BarraClose

ldFecIni = CDate(psFecIni)
ldFecFin = CDate(psFecFin)
nDias = ldFecFin - ldFecIni + 1

' Inicializamos variables
ReDim laSdoEnc(1 To lnNroAgencias, 1 To 3, 1 To nDias)  'Saldos de Encajes
ReDim laTotEnc(1 To lnNroAgencias, 1 To 3)              'Totales Enc por Agencia

ReDim laSdoEncInt(1 To lnNroAgencias + 1, 1 To 3, 1 To nDias) 'Saldos de Encajes para interes
ReDim laTotEncInt(1 To lnNroAgencias + 1, 1 To 3)             'Totales Enc por Agencia para interes

ReDim laSdoBil(1 To lnNroAgencias + 2, 1 To nDias)      'Saldo de Billetes por Agencia
ReDim laTotBil(1 To lnNroAgencias + 2)                  'Totales Billete por Agencia
ReDim laProvis(1 To 8, 1 To nDias)                       'Para Planilla de Provision
ReDim laTotPro(1 To 8)                                   'Totales por Concepto de Provision
For K = 1 To lnNroAgencias
   For I = 1 To nDias
      For J = 1 To 3
         laSdoEnc(K, J, I) = "0"
         laSdoEncInt(K, J, I) = "0"
      Next J
   Next I
   laTotBil(K) = 0
Next K
laTotBil(lnNroAgencias + 1) = 0
laTotBil(lnNroAgencias + 2) = 0

For K = 1 To lnNroAgencias
   For I = 1 To 3
      laTotEnc(K, I) = 0
      laTotEncInt(K, I) = 0
   Next
   For I = 1 To nDias
      laSdoEnc(K, 1, I) = CDate(psFecIni) + I - 1
      laSdoBil(K, I) = "0"
   Next
Next
For I = 1 To nDias
   For J = 1 To 8
      laProvis(J, I) = "0"
      laTotPro(J) = 0
   Next
Next
If lnTipo = 1 Then
   
   'Cuenta Contable de efectivo para BilletajeCaja
   CalculaBilleteCaja lsMoneda, lsCodCtaCaja, ldFecIni, ldFecFin
   
   'Cuenta Contable de efectivo para Cuenta de BCR y objeto para cuenta de BCR
   CalculaSaldoBCR lsMoneda, lsCodCtaBanc, ldFecIni, ldFecFin

End If
lnContAge = 0
Do While Not rs.EOF
   RaiseEvent BarraShow(1)
   RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", Trim(rs!Descripcion), "Procesando..", vbBlue)
   lnContAge = lnContAge + 1
   If rs!Nivel = 2 Then
      
      CalculaSdoEnc pnBitCentral_, lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
      
      If lnTipo = 1 Then
         
         CalculaBilleteAge lsMoneda, rs!Codigo, lsCodCtaCAge, ldFecIni, ldFecFin, lnContAge
      
      End If
      'CalculaSdoEncInteres lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
      CalculaSdoEncObligaciones lsMoneda, rs!Codigo, ldFecIni, ldFecFin, lnContAge
   End If
   RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", Trim(rs!Descripcion), "Procesando..", vbBlue)
   rs.MoveNext
Loop

'Calculo Obligaciones Inmediatas Para Caja

dTempo = ldFecIni
Do While Format(ldFecFin, "YYYYmmdd") >= Format(dTempo, "YYYYmmdd")
    lnDiaMes = dTempo - ldFecIni + 1
    nTotal = 0
    nTotal = SaldoObligInmediatas(1, "76" & lsMoneda & "201", dTempo, ldFecFin, "99")
    laSdoEncInt(UBound(laSdoEncInt, 1), 2, lnDiaMes) = nTotal
    laTotEncInt(UBound(laSdoEncInt, 1), 1) = Val(laTotEncInt(UBound(laSdoEncInt, 1), 1)) + laSdoEncInt(UBound(laSdoEncInt, 1), 2, lnDiaMes)
    dTempo = dTempo + 1
Loop

If lsMoneda = "1" Then
   RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "CALCULO DE CAJA PROMEDIO", "Procesando..", vbBlue)
   
   'En soles es el promedio del mes pasado
   CalculaBilleteTotCajaPromedio lsMoneda, lsCodCtaCaja, ldFecIni
   
   RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "CALCULO DE CAJA PROMEDIO", "Procesando..", vbBlue)
End If

Dim lsImpre As String

If lnTipo = 1 Then
    ImprimePlanillaAhorroPF 2, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
    
    ImprimePlanillaAhorroPF 3, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
    
    ImprimePlanillaObligaciones 4, lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
    
    If lnTipo = 1 Then
        ImprimePlanillaBilletaje lsMoneda, xlAplicacion, xlLibro, xlHoja1, psOpeCod
        
        CalculaTotProvis
        ImprimePlanillaProvision lsMoneda, Format(ldFecFin, gsFormatoFechaView), xlAplicacion, xlLibro, xlHoja1, psOpeCod
    End If
    ImprimeConsolidaSdoEnc = ""
   
'Quitado
'   lsImpre = CON & PrtSdoEnc(2, lsMoneda, xlAplicacion, xlLibro, xlHoja1) & PrtSdoEnc(3, lsMoneda, xlAplicacion, xlLibro, xlHoja1) & PrtSdoEncInt(4, lsMoneda, xlAplicacion, xlLibro, xlHoja1)
'   lsImpre = lsImpre & PrtSdoBillete(lsMoneda, xlAplicacion, xlLibro, xlHoja1)
'   CalculaTotProvis
   
'   lsImpre = lsImpre & PrtPlanillaProvision(lsMoneda, Format(ldFecFin, gsFormatoFechaView), xlAplicacion, xlLibro, xlHoja1) & COFF
'Fin Quitado

End If
RaiseEvent BarraClose
'ImprimeConsolidaSdoEnc = lsImpre

End Function

'Inicio Agregado
Private Sub ImprimePlanillaAhorroPF(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    Dim cNameHoja As String
    
    If nTipo = 2 Then
        cNameHoja = "AhorroC_Capital" 'ok
    ElseIf nTipo = 3 Then
        cNameHoja = "PlazoF_Capital" 'ok
    End If
    ExcelAddHoja cNameHoja, xlLibro, xlHoja1
        
    'ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
    
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
     
    lnLinPag = 65
    lnCarLin = UBound(laSdoEnc, 1) * 20   '  180
    If lnTipo = 1 Then
      lsTitRp1 = "E F E C T O S   D E   E N C A J E (CAPITAL)"
    Else
      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
    End If
    If nTipo = 2 Then
      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
    Else
      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
    
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
    
'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    xlHoja1.Cells(6, I + 1) = "TOTAL"
    
    lnLinPag = 8
    For I = LBound(laSdoEnc, 3) To UBound(laSdoEnc, 3)
      
        xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, I), gsFormatoFechaView)
       
        nTotAge = 0
        For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
            
            xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoEnc(K, nTipo, I))
            
            nTotAge = nTotAge + nVal(laSdoEnc(K, nTipo, I))
        Next
        
        laProvis(1, I) = laSdoEnc(1, 1, I)
        laProvis(nTipo, I) = nTotAge
        laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
        
        'Totales por Dia
        xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
        
    Next I

    nTotAge = 0
    For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
        'Totales Por Agencia
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotEnc(K, nTipo)
        nTotAge = nTotAge + laTotEnc(K, nTipo)
    Next
    
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge

    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    
End Sub
Private Sub ImprimePlanillaObligaciones(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    Dim cNameHoja As String
    
    If nTipo = 4 Then
        cNameHoja = "Oblig_Inmediatas" ' ok
    End If
    
    ExcelAddHoja cNameHoja, xlLibro, xlHoja1
    
    'ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
    
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
    lnLinPag = 65
    lnCarLin = UBound(laSdoEncInt, 1) * 20   '  180
    If lnTipo = 1 Then
      lsTitRp1 = "E F E C T O S   D E   E N C A J E (INTERES)"
    Else
      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
    End If
    If nTipo = 2 Then
      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
    Else
      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
    End If
    
    If nTipo = 4 Then
        lsTitRp1 = "O B L I G A C I O N E S   I N M E D I A T A S "
    End If
    
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
                
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
    
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True

'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    If lnTipo = 1 Then
        xlHoja1.Cells(6, I + 1) = "CAJA GENERAL"
        xlHoja1.Cells(6, I + 2) = "TOTAL"
    Else
        xlHoja1.Cells(6, I + 1) = "TOTAL"
    End If
    
    lnLinPag = 8
    For I = LBound(laSdoEncInt, 3) To UBound(laSdoEncInt, 3)
        xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, I), gsFormatoFechaView)
      
        nTotAge = 0
        For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
    
            xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoEncInt(K, 2, I))
            
            nTotAge = nTotAge + nVal(laSdoEncInt(K, 2, I))
        Next
        
        'adiciona la provision para el total en la planilla provisional de encaje
        laProvis(nTipo, I) = nTotAge
        laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
                
        'Totales por Dia
        xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge

    Next I
    
    nTotAge = 0
    
    For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
        'Totales Por Agencia
        
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotEncInt(K, 1)
        
        nTotAge = nTotAge + laTotEncInt(K, 1)
    Next
 
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge

    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
     
End Sub

Private Sub ImprimePlanillaBilletaje(psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    
    ExcelAddHoja "Billetaje", xlLibro, xlHoja1
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
     
    lnLinPag = 65
    lnCarLin = (UBound(laSdoEnc, 1) * 20) + 18   '198
    If psMoneda = "1" Then
      lsTitRp1 = "C A J A    M O N E D A    N A C I O N A L"
    Else
      lsTitRp1 = "C A J A    M O N E D A    E X T R A N J E R A"
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
     
'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
    
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
    
    xlHoja1.Cells(6, I + 1) = "CAJA GENERAL"
    xlHoja1.Cells(6, I + 2) = "TOTAL"
        
    lnLinPag = 8
    For I = LBound(laSdoBil, 2) To UBound(laSdoBil, 2)
        xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoBil(UBound(laSdoBil, 1), I), gsFormatoFechaView)
        
        nTotAge = 0
        For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
      
            xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoBil(K, I))
                
            nTotAge = nTotAge + nVal(laSdoBil(K, I))
        Next
        
        If psMoneda = "2" Then
            laProvis(6, I) = nTotAge
            laTotPro(6) = laTotPro(6) + nTotAge
        End If
        
        'Totales por Dia
        
        xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
        
    Next I
    xlHoja1.Cells(I + lnLinPag, 1) = "Totales"
    
    nTotAge = 0
    For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
        'Totales por Agencia
        
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotBil(K)
        nTotAge = nTotAge + laTotBil(K)
    Next
    
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    
    Dim oEst As New NEstadisticas
    If Day(ldFecFin + 1) = 1 Then  'Es ultimo dia del mes
       oEst.EliminaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda
       oEst.InsertaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda, Round(nTotAge / Day(ldFecFin), 2)
    End If
    Set oEst = Nothing
     
End Sub

Private Sub ImprimePlanillaProvision(psMoneda As String, lsFecha As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, ppsopecod As String)
    Dim I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    
    'Variables Calculo de Encaje
    Dim R             As ADODB.Recordset
    Dim MatTemp()     As String
    Dim MatTemp2()    As String
    Dim MatTemp3()    As String
    Dim Total1        As Double
    Dim Total2        As Double
    Dim lnTotalEncaje As Double
    Dim pdFecIni As Date
    Dim oPara As New NEncajeBCR
    Dim lnLinPag As Long

    pdFecIni = Format("01" & Mid(lsFecha, 3), "MM/dd/YYYY")
    
    
'    'Variables Calculo Encaje M.E.
'    Dim lnSEncProm As Double, lnSEncMarg As Double
'    Dim lnTasaBase As Double, lnTasaEncMarg As Double
'    Dim lnEncMarg As Double, lnEncCalc As Double
'    Dim lnTObligEnc As Double, lnEncExig As Double
'    Dim lnTasaEncLegMN As Double
'    Dim lnEncLeg As Double
'    Dim lnLinPag As Integer
'    Dim lnEncMarColoc As Double
    
    ExcelAddHoja "Planilla", xlLibro, xlHoja1
    xlHoja1.PageSetup.Zoom = 80
    
    lnLinPag = 65
     
    Dim oBala As NBalanceCont
    Set oBala = New NBalanceCont
    
    'lnTasaEncLegMN = oBala.GetUtilidadAcumulada("P", -1, Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
    If psMoneda = "1" Then
      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA NACIONAL"
    Else
      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA EXTRANJERA"
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 8) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 4)).MergeCells = True
    
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).Font.Bold = True

    xlHoja1.Cells(6, 1) = "FECHA"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 1)).MergeCells = True
    
    xlHoja1.Cells(6, 2) = "DEPOSITOS Y OBLIGACIONES  SUJETOS DE ENCAJE": xlHoja1.Cells(7, 2) = "AHORRO": xlHoja1.Cells(7, 3) = "PLAZO FIJO": xlHoja1.Cells(7, 4) = "OBLIG. INMED.": xlHoja1.Cells(7, 5) = "TOTAL"
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).MergeCells = True
    
    xlHoja1.Cells(6, 6) = "EFECTIVO Y DEPOSITOS DE ENCAJE":  xlHoja1.Cells(7, 6) = "CAJA": xlHoja1.Cells(7, 7) = "DEPOSITOS BCR":     xlHoja1.Cells(7, 8) = "TOTAL"
    xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    lnLinPag = 9
    For I = LBound(laProvis, 2) To UBound(laProvis, 2) 'Recorro el Numero de Dias
            
      'Dia
      xlHoja1.Cells(I + lnLinPag - 1, 1) = "'" & Format(laProvis(1, I), gsFormatoFechaView)
      For K = 2 To 8
          '2 = Ahorro; 3 = PF; 4 = ObligInm; 5 = Total; 6=Caja; 7= Depositos BCR; 8 =Total
          xlHoja1.Cells(I + lnLinPag - 1, K) = nVal(laProvis(K, I))
      Next
    Next I
    xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(I + lnLinPag - 2, 8)).BorderAround xlContinuous, xlMedium
    
    'Total de Todos los Dias
    For K = 2 To 8
        xlHoja1.Cells(I + lnLinPag - 1, K) = laTotPro(K)
    Next
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag - 1, 1), xlHoja1.Cells(I + lnLinPag - 1, 8)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(I + lnLinPag - 1, 8)).NumberFormat = "#,##0.00"
    
    'Imprimimos Cómputo del Encaje
    
    xlHoja1.Cells(I + lnLinPag + 1, 2) = "DIAS"
    xlHoja1.Cells(I + lnLinPag + 1, 3) = "ENCAJE CMAC"
    xlHoja1.Cells(I + lnLinPag + 1, 4) = "ENCAJE LEGAL"
    xlHoja1.Cells(I + lnLinPag + 1, 5) = "EXCEDENTE/DEFICIT"
    
    'Calculo del Encaje
    '''''''''''''''''''
'''''    If psMoneda = "1" Then
'''''      lnEncLeg = Round((laTotPro(5)) * lnTasaEncLegMN / 100, 2) & "   "
'''''    Else
'''''      lnTObligEnc = oBala.GetUtilidadAcumulada("T", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''
'''''      If Format(lsFecha, "YYYYMMdd") >= "20040401" And Format(lsFecha, "YYYYMMdd") <= "20041031" Then
'''''        lnSEncProm = Round(lnTObligEnc / 29, 2)
'''''      ElseIf Format(lsFecha, "YYYYMMdd") >= "20041101" Then
'''''        lnSEncProm = Round(lnTObligEnc / 30, 2)
'''''      Else
'''''        lnSEncProm = Round(lnTObligEnc / 31, 2)
'''''      End If
'''''
'''''      lnSEncMarg = Round(((laTotPro(5) - lnTotalColocFSD) / nDias) - lnSEncProm, 2)
'''''      lnEncExig = oBala.GetUtilidadAcumulada("E", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'''''      lnTasaBase = Round(lnEncExig / lnTObligEnc, 6)
'''''      If CDate(lsFecha) >= CDate("01/09/2000") And CDate(lsFecha) <= CDate("30/11/2000") Then
'''''            lnTasaBase = lnTasaBase - 0.03
'''''      End If
'''''
'''''      lnTasaEncMarg = oBala.GetUtilidadAcumulada("P", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True) / 100
'''''      'lnTasaEncMarg = 0.2
'''''
'''''      If lnSEncMarg > 0 Then
'''''         lnEncMarg = Round(lnSEncMarg * lnTasaEncMarg, 2)
'''''         lnEncMarColoc = Round(((lnTotalColocFSD) * lnTasaEncMarg), 2)
'''''         lnEncCalc = Round(lnSEncProm * lnTasaBase, 2)
'''''      Else
'''''         lnEncMarg = 0
'''''         lnEncMarColoc = Round(((lnTotalColocFSD) * lnTasaEncMarg), 2)
'''''         lnEncCalc = Round(Round(((laTotPro(5) - lnTotalColocFSD) / nDias), 2) * lnTasaBase, 2)
'''''      End If
'''''
'''''      lnEncLeg = Round((lnEncCalc + lnEncMarg) * nDias, 2) + lnEncMarColoc
'''''
'''''    End If
'''''
'''''    'Dias
'''''    xlHoja1.Cells(i + lnLinPag + 2, 2) = nDias
'''''    cCodigo1 = "P_DIAS"
'''''    cCodigo2 = ""
'''''    cCodigo3 = ""
'''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Trim(Str(nDias))
'''''
'''''    'Encaje CMAC
'''''    xlHoja1.Cells(i + lnLinPag + 2, 3) = laTotPro(8)
'''''    cCodigo1 = "P_ENCAJECMAC"
'''''    cCodigo2 = ""
'''''    cCodigo3 = ""
'''''
'''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(laTotPro(8), "0.00")
'''''
'''''    'Encaje Legal
'''''    xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
'''''    cCodigo1 = "P_ENCAJELEGAL"
'''''    cCodigo2 = ""
'''''    cCodigo3 = ""
'''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(Round(lnEncLeg, 2), "0.00")
'''''
'''''    'Excedente/Deficit ==> Es el encaje_Acumula que se grababa antes
'''''    xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
'''''    cCodigo1 = "P_EXCEDENTEDEFICIT"
'''''    cCodigo2 = ""
'''''    cCodigo3 = ""
'''''    GrabaCGRptConsolidado ppsopecod, ldFecFin, cCodigo1, cCodigo2, cCodigo3, Val(psMoneda), Format(Round(laTotPro(8) - lnEncLeg, 2), "0.00")
        
    If psMoneda = "2" Then
        Set R = oPara.GetParametroEncaje(pdFecIni)
        ReDim MatTemp(UBound(laProvis, 2))
        ReDim MatTemp2(UBound(laProvis, 2))
        Total1 = 0
        Total2 = 0
        R.Find "nCodigo=2"
      
        'Hallo Exceso diario EncajeActual - Base
        For I = 1 To UBound(laProvis, 2)
            MatTemp(I) = nVal(laProvis(5, I)) - R!nValor
        Next I
      
        Total1 = R!nValor
      
        R.Find "nCodigo=3"
        'Hallo El Basico : Base * TasaBase (21.6478)%
        For I = 1 To UBound(laProvis, 2)
            MatTemp2(I) = Total1 * R!nValor
        Next I
        Total1 = 0#
      
        R.Find "nCodigo=4"
        
        'Encaje Marginal: Marginal por el Exeso * 20%
        For I = 1 To UBound(MatTemp)
            If MatTemp(I) > 0 Then
                MatTemp(I) = CDbl(MatTemp(I)) * R!nValor
            Else
                MatTemp(I) = 0
            End If
        Next I
      
        'Total Encaje : En Basico mas Marginal
        For I = 1 To UBound(MatTemp)
            MatTemp2(I) = CDbl(MatTemp(I)) + CDbl(MatTemp2(I))
        Next I
      
        For I = 1 To UBound(MatTemp)
            Total1 = Total1 + CDbl(MatTemp2(I))
        Next I
        Total1 = CDbl(Format(Total1, "#0.00"))
          
        R.Close
        Set R = Nothing
    Else
        Set R = oPara.GetParametroEncaje(pdFecIni)
        R.Find "nCodigo=1"
        Total2 = R!nValor
        R.Close
        Set R = Nothing
    End If
    
    If psMoneda = "1" Then
       lnTotalEncaje = CCur(laTotPro(5)) * Total2
   Else
       lnTotalEncaje = Total1
   End If
    
    'Dias
    xlHoja1.Cells(I + lnLinPag + 2, 2) = nDias
    'Encaje CMAC
    xlHoja1.Cells(I + lnLinPag + 2, 3) = laTotPro(8)
    'Encaje Legal
    'xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
    xlHoja1.Cells(I + lnLinPag + 2, 4) = Round(lnTotalEncaje, 2)
    'Excedente
    'xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
    xlHoja1.Cells(I + lnLinPag + 2, 5) = Round(laTotPro(8) - lnTotalEncaje, 2)
    
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 2, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).NumberFormat = "#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 3), xlHoja1.Cells(I + lnLinPag + 2, 4)).ColumnWidth = 14
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 5), xlHoja1.Cells(I + lnLinPag + 2, 5)).ColumnWidth = 18
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    
    Dim oEst As New NEstadisticas
    'oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
    'oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnEncLeg, 2)
    oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
    oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnTotalEncaje, 2)

    Set oEst = Nothing
     
     
    Set oBala = Nothing
     
End Sub



'Fin Agregado

Private Sub CalculaBilleteAge(psMoneda As String, psAgeCod As String, psCodCta As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
   Dim lsQryEnc As String
   Dim rsQryEnc As New ADODB.Recordset
   Dim lnDiaMes As Integer
   Dim sCtaCod As String
   Dim lnSaldo As Currency
   sCtaCod = psCodCta
   Dim oConect As New DConecta
   oConect.AbreConexion
   
   lsQryEnc = "SELECT Fec.dFecha, nMonto = ISNULL(( SELECT SUM(nMonto) "
   lsQryEnc = lsQryEnc & "                  FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro "
   lsQryEnc = lsQryEnc & "                  WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and "
   lsQryEnc = lsQryEnc & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo me1 ON me1.nMovNro = m1.nMovNro WHERE me1.cEfectivoCod Like '" & psMoneda & "%' and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and m1.cMovNro Like '_________________" & psAgeCod & "%' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) and m1.nMovEstado = " & gMovEstContabNoContable & " "
   lsQryEnc = lsQryEnc & "                                and m1.nMovFlag = " & gMovFlagVigente & " ) and "
   lsQryEnc = lsQryEnc & "       mue.cEfectivoCod Like '" & psMoneda & "%' and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and "
   lsQryEnc = lsQryEnc & "       m.cMovNro Like '_________________" & psAgeCod & "%'),0)  FROM FechaTmp('" & Format(pdFecIni, gsFormatoFecha) & "')  Fec  "

   lsQryEnc = lsQryEnc & " WHERE Fec.dFecha >='" & Format(pdFecIni, "YYYYMMdd") & "' "
   lsQryEnc = lsQryEnc & " And Fec.dFecha <='" & Format(pdFecFin, "YYYYmmdd") & "' "
   
   lsQryEnc = lsQryEnc & " ORDER BY Fec.dFecha  "
   Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
    lnSaldo = 0
    lnDiaMes = 0
    Do While Not rsQryEnc.EOF
       lnDiaMes = lnDiaMes + 1
       lnSaldo = rsQryEnc!nMonto
       laSdoBil(pnContAge, lnDiaMes) = lnSaldo
       laTotBil(pnContAge) = laTotBil(pnContAge) + Val(laSdoBil(pnContAge, lnDiaMes))
       rsQryEnc.MoveNext
    Loop
   oConect.CierraConexion
   Set oConect = Nothing
End Sub

Private Sub CalculaBilleteAgePromedio(psMoneda As String, psAgeCod As String, psCodCta As String, pdFecIni As Date, pdFecFin As Date)
   Dim lsQryEnc As String
   Dim rsQryEnc As New ADODB.Recordset
   Dim lnDiaMes As Integer
   Dim lnAgeCod As Integer
   Dim sCtaCod As String
   Dim lnSaldo As Currency
   Dim lnDiaIni As Integer
   sCtaCod = psCodCta
   lnAgeCod = Val(psAgeCod)
   Dim oConect As New DConecta
   oConect.AbreConexion
'Centralizado
   lsQryEnc = "SELECT Fec.dFecha, SubString(m.cMovNro,18,2) cAgeCod, SUM(nMonto) nMonto " _
         & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
         & "       FechaTmp('" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "')  Fec " _
         & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Fec.dFecha BETWEEN '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "' and '" & Format(ldFecIni - 1, gsFormatoFecha) & "' and " _
         & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 WHERE m1.cOpeCod = m.cOpeCod and SubString(m1.cMovNro,18,2) =  '" & psAgeCod & "' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
         & "       LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
         & "       SubString(m.cMovNro,18,2) = '" & psAgeCod & "' " _
         & "GROUP BY Fec.dFecha, SubString(m.cMovNro,18,2) " _
         & "ORDER BY Fec.dFecha"
         
   Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
    lnSaldo = 0
    lnDiaMes = 0
    lnDiaIni = 1
    If Not rsQryEnc.EOF Then
        lnDiaIni = Day(rsQryEnc!dFecha)
    End If
    Do While Not rsQryEnc.EOF
       lnDiaMes = lnDiaMes + 1
       lnSaldo = lnSaldo + rsQryEnc!nMonto
'       lnSaldo = lnSaldo + nval(IIf(psMoneda = "1", rsQryEnc!nMonto, rsQryEnc!nCtaSaldoImporteME))
       rsQryEnc.MoveNext
    Loop
    If lnDiaMes > 0 Then
        lnSaldo = Round(lnSaldo / lnDiaMes, 2)
        For lnDiaMes = lnDiaIni To UBound(laSdoBil, 2)
            'laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
            'laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
            'laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
            laProvis(6, lnDiaMes) = laProvis(6, lnDiaMes) + lnSaldo
            laTotPro(6) = laTotPro(6) + lnSaldo
        Next
    End If
   oConect.CierraConexion
   Set oConect = Nothing
End Sub

'Private Sub CalculaBilleteTotCajaPromedio(psMoneda As String, psCtaCod As String, pdFecIni As Date)
'Dim lsQryEnc As String
'Dim rsQryEnc As ADODB.Recordset
'Dim lnDiaMes As Integer
'Dim lnSaldo As Currency
'Dim lnTotal As Currency
'
'    lnSaldo = 0
'    lnDiaMes = 0
'
'    Dim oEst As New NEstadisticas
'    lsQryEnc = oEst.GetImporteEstadAnexos(pdFecIni - 1, "CAJA_PROMEDIO", psMoneda)
'    If lsQryEnc = "0" Then
'
'    'Calcula el Promedio en funcion al Total Caja / Nro Dias
'      lsQryEnc = "SELECT ROUND(SUM(nMontoMN)/Count(*),2) nMontoMN, ROUND(SUM(nMontoME)/Count(*),2) nMontoME " _
'            & "FROM (SELECT dFecha, SUM(nMontoMN) nMontoMN, SUM(nMontoME) nMontoME " _
'            & "      FROM ( SELECT Fec.dFecha, SUM(nMonto) nMontoMN, SUM(nMonto) nMontoME " _
'            & "        FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
'            & "            FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "')  Fec " _
'            & "        WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Datediff(m,Fec.dFecha, '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "') = 0 and " _
'            & "            LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo me1 ON me1.nMovNro = m1.nMovNro WHERE m1.nMovEstado = " & gMovEstContabNoContable & " and m1.nMovFlag = " & gMovFlagVigente & " and LEFT(me1.cEfectivoCod,1) = " & psMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  SubString(m.cMovNro,18,2) and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
'            & "            LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
'            & "        GROUP BY Fec.dFecha UNION ALL " _
'            & "        SELECT Fec.dFecha, SUM(nCtaSaldoImporte) nMontoMN, SUM(nCtaSaldoImporteME) nMontoME  " _
'            & "         FROM CTASALDO CS1, FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') Fec " _
'            & "         WHERE cs1.cCtaContCod = '" & psCtaCod & "' and datediff(m,Fec.dFecha,'" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') = 0 and " _
'            & "             cs1.dCtaSaldoFecha = ( " _
'            & "                 SELECT  MAX(dCtaSaldoFecha) " _
'            & "                 FROM    CtaSaldo CS " _
'            & "                 WHERE   cs1.cCtaContCod = CS.cCtaContCod AND CS.dCtaSaldoFecha <= Fec.dFecha ) " _
'            & "         GROUP BY Fec.dFecha " _
'            & "      ) Caja GROUP BY dFecha ) C "
'
'       Dim oConect As New DConecta
'       oConect.AbreConexion
'       Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
'        If Not rsQryEnc.EOF Then
'           lnDiaMes = lnDiaMes + 1
'           lnSaldo = IIf(IsNull(rsQryEnc!NMONTOMN), 0, rsQryEnc!NMONTOMN)
'           oEst.InsertaEstadAnexos pdFecIni - 1, "CAJA_PROMEDIO", psMoneda, Str(lnSaldo)
'           rsQryEnc.MoveNext
'        End If
'    Else
'        lnDiaMes = lnDiaMes + 1
'        lnSaldo = nVal(lsQryEnc)
'    End If
'    If lnDiaMes > 0 Then
'        For lnDiaMes = 1 To UBound(laSdoBil, 2)
'            laProvis(6, lnDiaMes) = laProvis(6, lnDiaMes) + lnSaldo
'            laTotPro(6) = laTotPro(6) + lnSaldo
'        Next
'    End If
'   oConect.CierraConexion
'   Set oConect = Nothing
'End Sub

Private Sub CalculaBilleteTotCajaPromedio(psMoneda As String, psCtaCod As String, pdFecIni As Date)
Dim lsQryEnc As String
Dim rsQryEnc As ADODB.Recordset
Dim lnDiaMes As Integer
Dim lnSaldo As Currency
Dim lnTotal As Currency

    lnSaldo = 0
    lnDiaMes = 0

    Dim oEst As New NEstadisticas
    lsQryEnc = oEst.GetImporteEstadAnexos(pdFecIni - 1, "CAJA_PROMEDIO", psMoneda)
    If lsQryEnc = "0" Then
           
    'Calcula el Promedio en funcion al Total Caja / Nro Dias
      lsQryEnc = "SELECT ROUND(SUM(nMontoMN)/Count(*),2) nMontoMN, ROUND(SUM(nMontoME)/Count(*),2) nMontoME " _
            & "FROM (SELECT dFecha, SUM(nMontoMN) nMontoMN, SUM(nMontoME) nMontoME " _
            & "      FROM ( SELECT Fec.dFecha, SUM(nMonto) nMontoMN, SUM(nMonto) nMontoME " _
            & "        FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
            & "            FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "')  Fec " _
            & "        WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Datediff(m,Fec.dFecha, '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "') = 0 and " _
            & "            LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo me1 ON me1.nMovNro = m1.nMovNro WHERE m1.nMovEstado = " & gMovEstContabNoContable & " and m1.nMovFlag = " & gMovFlagVigente & " and LEFT(me1.cEfectivoCod,1) = " & psMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  SubString(m.cMovNro,18,2) and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
            & "            LEFT(mue.cEfectivoCod,1) = " & psMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') " _
            & "        GROUP BY Fec.dFecha UNION ALL " _
            & "        SELECT Fec.dFecha, SUM(nCtaSaldoImporte) nMontoMN, SUM(nCtaSaldoImporteME) nMontoME  " _
            & "         FROM CTASALDO CS1, FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') Fec " _
            & "         WHERE cs1.cCtaContCod = '" & psCtaCod & "' and datediff(m,Fec.dFecha,'" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "') = 0 and " _
            & "             cs1.dCtaSaldoFecha = ( " _
            & "                 SELECT  MAX(dCtaSaldoFecha) " _
            & "                 FROM    CtaSaldo CS " _
            & "                 WHERE   cs1.cCtaContCod = CS.cCtaContCod AND CS.dCtaSaldoFecha <= Fec.dFecha ) " _
            & "         GROUP BY Fec.dFecha " _
            & "      ) Caja GROUP BY dFecha ) C "
    
    
      lsQryEnc = " SELECT ROUND(SUM(nMontoMN)/Count(*),2) nMontoMN, ROUND(SUM(nMontoME)/Count(*),2) nMontoME " _
               & " FROM (" _
               & "   Select dFecha, dbo.GetSaldoCta(dFecha,'11110101',1) + dbo.GetSaldoCtaAcumulado(dFecha,'11110[23]%',1) nMontoMN, " _
               & "   Convert(decimal(20,2),(dbo.GetSaldoCta(dFecha,'11210101',1) + dbo.GetSaldoCtaAcumulado(dFecha,'11210[23]%',1)) / dbo.GetTipoCambioFinMes(1,dfecha)) nMontoME" _
               & "   From dbo.FechaTmp('" & Format(DateAdd("m", -1, pdFecIni), gsFormatoFecha) & "')" _
               & " ) C "
    
    
       Dim oConect As New DConecta
       oConect.AbreConexion
       Set rsQryEnc = oConect.CargaRecordSet(lsQryEnc)
        If Not rsQryEnc.EOF Then
           lnDiaMes = lnDiaMes + 1
           lnSaldo = IIf(IsNull(rsQryEnc!NMONTOMN), 0, rsQryEnc!NMONTOMN)
           oEst.InsertaEstadAnexos pdFecIni - 1, "CAJA_PROMEDIO", psMoneda, Str(lnSaldo)
           rsQryEnc.MoveNext
        End If
    Else
        lnDiaMes = lnDiaMes + 1
        lnSaldo = nVal(lsQryEnc)
    End If
    If lnDiaMes > 0 Then
        For lnDiaMes = 1 To UBound(laSdoBil, 2)
            laProvis(6, lnDiaMes) = laProvis(6, lnDiaMes) + lnSaldo
            laTotPro(6) = laTotPro(6) + lnSaldo
        Next
    End If
   oConect.CierraConexion
   Set oConect = Nothing
End Sub


Private Sub CalculaBilleteCaja(psMoneda As String, psCodCtaCaja As String, pdFecIni As Date, pdFecFin As Date)
    Dim lsQryEnc As String
    Dim rsQryEnc As New ADODB.Recordset
    Dim lnDiaMes As Integer
    Dim lnAgeCod As Integer
    Dim lnSaldo  As Currency
    
    Dim oSdo As New NCtasaldo
    RaiseEvent BarraShow(1)
    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaCaja, pdFecIni, pdFecFin)
    lnSaldo = 0
    lnDiaMes = 0
    Do While Not rsQryEnc.EOF
      lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
      lnSaldo = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtasaldoImporteME))
      laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
      laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
      laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
      rsQryEnc.MoveNext
    Loop
    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
    RaiseEvent BarraClose
    Set oSdo = New NCtasaldo
End Sub

Private Sub CalculaBilleteCajaPromedio(psMoneda As String, psCodCtaCaja As String, pdFecIni As Date, pdFecFin As Date)
    Dim lsQryEnc As String
    Dim rsQryEnc As New ADODB.Recordset
    Dim lnDiaMes As Integer
    Dim lnAgeCod As Integer
    Dim lnSaldo  As Currency
    Dim lnTotal As Currency
    
    Dim oSdo As New NCtasaldo
    RaiseEvent BarraShow(1)
    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaCaja, DateAdd("m", -1, ldFecIni), ldFecIni - 1)
    lnSaldo = 0
    lnDiaMes = 0
    Do While Not rsQryEnc.EOF
       lnDiaMes = lnDiaMes + 1
       lnSaldo = lnSaldo + nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtasaldoImporteME))
       rsQryEnc.MoveNext
    Loop
    If lnDiaMes > 0 Then
        lnSaldo = Round(lnSaldo / lnDiaMes, 2)
        lnTotal = 0
        For lnDiaMes = 1 To UBound(laSdoBil, 2)
'            laSdoBil(UBound(laSdoBil, 1), lnDiaMes) = pdFecIni + lnDiaMes - 1
'            laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes) = lnSaldo
'            laTotBil(UBound(laSdoBil, 1) - 1) = laTotBil(UBound(laSdoBil, 1) - 1) + Val(laSdoBil(UBound(laSdoBil, 1) - 1, lnDiaMes))
            lnTotal = lnTotal + lnSaldo
            laProvis(6, lnDiaMes) = lnSaldo
        Next
        laTotPro(6) = laTotPro(6) + lnTotal
    End If
    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "BILLETAJE CAJA GENERAL", "Procesando...", vbBlue)
    RaiseEvent BarraClose
    Set oSdo = New NCtasaldo
End Sub

Private Sub CalculaSaldoBCR(psMoneda As String, psCodCtaBCR As String, pdFecIni As Date, pdFecFin As Date)
    Dim lsQryEnc As String
    Dim rsQryEnc As New ADODB.Recordset
    Dim lnDiaMes As Integer
    Dim lnSaldo  As Currency

    Dim oSdo As New NCtasaldo
    RaiseEvent BarraShow(1)
    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS BANCO BCR", "Procesando...", vbBlue)
    Set rsQryEnc = oSdo.GetCtaSaldoRango(psCodCtaBCR & "%", pdFecIni, pdFecFin)
    Do While Not rsQryEnc.EOF
       lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
       laProvis(1, lnDiaMes) = Format(rsQryEnc!dFecha, gsFormatoFechaView)
       laProvis(7, lnDiaMes) = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtasaldoImporteME))
       rsQryEnc.MoveNext
    Loop
    For lnDiaMes = 1 To UBound(laProvis, 2)
      If lnDiaMes = 1 Then
         If Val(laProvis(7, 1)) = 0 Then
            lnSaldo = oSdo.GetCtaSaldo(psCodCtaBCR, Format(pdFecIni, gsFormatoFecha), psMoneda = "1")
            laProvis(1, lnDiaMes) = pdFecIni
            laProvis(7, lnDiaMes) = lnSaldo
         End If
      Else
         If Val(laProvis(7, lnDiaMes)) = 0 Then
            laProvis(1, lnDiaMes) = pdFecIni + lnDiaMes - 1
            laProvis(7, lnDiaMes) = laProvis(7, lnDiaMes - 1)
         End If
      End If
      laTotPro(7) = laTotPro(7) + Val(laProvis(7, lnDiaMes))
    Next
    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS BANCO BCR", "Procesando...", vbBlue)
    Set oSdo = New NCtasaldo
End Sub

Private Sub CalculaSdoEncObligaciones(psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)

    Dim dTempo As Date
    Dim lnDiaMes As Integer
    Dim Total As Double

    RaiseEvent BarraShow(1)
    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS OBLIGACIONES INMEDIATAS", "Procesando..", vbBlue)

    dTempo = pdFecIni
    
    Do While Format(pdFecFin, "YYYYmmdd") >= Format(dTempo, "YYYYmmdd")
        
        lnDiaMes = dTempo - pdFecIni + 1
        
        Total = SaldoObligInmediatas(1, "76" & psMoneda & "201", dTempo, pdFecFin, psAgeCod)
        laSdoEncInt(pnContAge, 1, lnDiaMes) = Format(dTempo, gsFormatoFechaView)
        laSdoEncInt(pnContAge, 2, lnDiaMes) = Total
        laTotEncInt(pnContAge, 1) = Val(laTotEncInt(pnContAge, 1)) + laSdoEncInt(pnContAge, 2, lnDiaMes)
       
        dTempo = dTempo + 1
       
    Loop

    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS OBLIGACIONES INMEDIATAS", "Procesando...", vbBlue)
End Sub

Public Function SaldoObligInmediatas(lnNroCol As Long, lsOpeCod As String, ByVal ldFecha As Date, ByVal pdFecFin As Date, ByVal cCodAge As String) As Currency
Dim rsC As New ADODB.Recordset
Dim oRep As New DRepCtaColumna

Dim oTC As New nTipoCambio
Dim pnTCF As Double
Dim pnTCFF As Double

pnTCF = Format(oTC.EmiteTipoCambio(pdFecFin, TCFijoMes), "#,##0.00###")
pnTCFF = Format(oTC.EmiteTipoCambio(pdFecFin + 1, TCFijoMes), "#,##0.00###")

Set rsC = oRep.GetRepColumnaCtaSaldo(lsOpeCod, lnNroCol, ldFecha, gbBitCentral, cCodAge)
SaldoObligInmediatas = 0
If Not RSVacio(rsC) Then
    If Mid(lsOpeCod, 3, 1) = "1" Then
        SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, rsC!Total)
    Else
        If pdFecFin = ldFecha Then
            If pnTCFF = 0 Then
                SaldoObligInmediatas = 0
            Else
                SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCFF, 2))
            End If
        Else
            If pnTCF = 0 Then
                SaldoObligInmediatas = 0
            Else
                SaldoObligInmediatas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCF, 2))
            End If
        End If
    End If
End If
Set oRep = Nothing
RSClose rsC
End Function

'Private Sub CalculaSdoEncInteres(psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
'    Dim lsQryEnc As String
'    Dim rsQryEnc As New ADODB.Recordset
'    Dim lnDiaMes As Integer
'    Dim ldFecha     As Date
'    Dim lnSaldo     As Currency
'    Dim i           As Integer
'    Dim lsCtaCodInt As String
'    Dim rsCta As ADODB.Recordset
'
'    ldFecha = ldFecIni
'    Dim oDOpe As DOperacion
'    Dim oSdo  As NCtasaldo
'    RaiseEvent BarraShow(1)
'    RaiseEvent BarraProgress(0, "CONSOLIDADO SALDOS ENCAJE", "SALDOS ENCAJE INTERESES", "Procesando..", vbBlue)
'    Set oDOpe = New DOperacion
'    Set oSdo = New NCtasaldo
'    'lsCtaCodInt = oDOpe.EmiteOpeCta(lsOpeCod, "D", "3")
'    Set rsCta = oDOpe.CargaOpeCtaUltimoNivel(lsOpeCod, "D", 3)
'    Do While Not rsCta.EOF
'        If Right(rsCta!cCtaContCod, 2) = psAgeCod Then
'            lsCtaCodInt = lsCtaCodInt & "'" & rsCta!cCtaContCod & "',"
'        End If
'        rsCta.MoveNext
'    Loop
'    If Not lsCtaCodInt = "" Then
'        lsCtaCodInt = Left(lsCtaCodInt, Len(lsCtaCodInt) - 1)
'    End If
'    RSClose rsCta
'
'    Set rsQryEnc = oSdo.GetCtaSaldoRango(lsCtaCodInt, pdFecIni, pdFecFin)
'    Do While Not rsQryEnc.EOF
'       lnDiaMes = rsQryEnc!dFecha - pdFecIni + 1
'       laSdoEncInt(pnContAge, 1, lnDiaMes) = Format(rsQryEnc!dFecha, gsFormatoFechaView)
'       laSdoEncInt(pnContAge, 2, lnDiaMes) = nVal(IIf(psMoneda = "1", rsQryEnc!nCtaSaldoImporte, rsQryEnc!nCtaSaldoImporteME))
'       laTotEncInt(pnContAge, 1) = Val(laTotEncInt(pnContAge, 1)) + laSdoEncInt(pnContAge, 2, lnDiaMes)
'       rsQryEnc.MoveNext
'    Loop
'    RaiseEvent BarraProgress(1, "CONSOLIDADO SALDOS ENCAJE", "SALDOS ENCAJE INTERESES", "Procesando...", vbBlue)
'    Set oSdo = Nothing
'    Set oDOpe = Nothing
'End Sub




Private Sub CalculaSdoEnc(pnBitCentral_ As Boolean, psMoneda As String, psAgeCod As String, pdFecIni As Date, pdFecFin As Date, pnContAge As Integer)
    Dim lsQryEnc As String
    Dim rsQryEnc As New ADODB.Recordset
    Dim lnDiaMes As Integer
    Dim pdFechaControl As Date
    
    Dim oConect As New DConecta
    Dim oEst As New NEstadisticas
    oConect.AbreConexion
    pdFechaControl = pdFecIni
    Set rsQryEnc = oEst.GetEstadisticaAhorro(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
    Do While Not rsQryEnc.EOF
        lnDiaMes = rsQryEnc!dEstadAC - pdFecIni + 1
        If Day(pdFechaControl) <> lnDiaMes Then
'          MsgBox "No se encontraron datos Estadisticos AC del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
        Else
            If pnBitCentral_ = True Then
                If lnTipo = 1 Then
                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) - nVal(rsQryEnc!nMonChqVal)
                Else
                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) + nVal(rsQryEnc!nSaldoCMAC) + nVal(rsQryEnc!nSaldCRAC) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
                End If
            Else
                If lnTipo = 1 Then
                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoAC) + IIf(pnContAge = 1, nPPD, 0) - nVal(rsQryEnc!nSaldoCMAC) - nVal(rsQryEnc!nSaldCRAC)
                Else
                  laSdoEnc(pnContAge, 2, lnDiaMes) = nVal(rsQryEnc!nSaldoAC) + IIf(pnContAge = 1, nPPD, 0) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
                End If
            End If
          
            laTotEnc(pnContAge, 2) = Val(laTotEnc(pnContAge, 2)) + Val(laSdoEnc(pnContAge, 2, lnDiaMes))
            'rsQryEnc.MoveNext
        End If
        rsQryEnc.MoveNext
        pdFechaControl = pdFechaControl + 1
    Loop
    RSClose rsQryEnc
    pdFechaControl = pdFecIni
    Set rsQryEnc = oEst.GetEstadisticaPlazoFijo(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
    Do While Not rsQryEnc.EOF
        lnDiaMes = rsQryEnc!dEstadPF - pdFecIni + 1
        If Day(pdFechaControl) <> lnDiaMes Then
'          MsgBox "No se encontraron datos Estadisticos PF del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
        Else
            If pnBitCentral_ = True Then
                If lnDiaMes = 31 Then
                    'MsgBox "sss"
                End If
                If lnTipo = 1 Then
                    'Maynas no suma el Interes Devengado
                    'laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) - nVal(rsQryEnc!nMonChqVal)  + nVal(rsQryEnc!nIntDev)
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) - nVal(rsQryEnc!nMonChqVal)
                Else
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoSinCmac) + nVal(rsQryEnc!nSaldoCMAC) + nVal(rsQryEnc!nSaldCRAC) + nVal(rsQryEnc!nMonChqVal) + nVal(rsQryEnc!nChqCMAC)
                End If
            Else
                If lnTipo = 1 Then
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoPF) - nVal(rsQryEnc!nSaldoCMAC) - nVal(rsQryEnc!nSaldCRAC)
                Else
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(rsQryEnc!nSaldoPF) + nVal(rsQryEnc!nMonChqVal)
                End If
            End If
            laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + Val(laSdoEnc(pnContAge, 3, lnDiaMes))
            'rsQryEnc.MoveNext
        End If
        rsQryEnc.MoveNext
        pdFechaControl = pdFechaControl + 1
        
    Loop
    RSClose rsQryEnc
        
    pdFechaControl = pdFecIni
    Set rsQryEnc = oEst.GetEstadisticaCTS(pnBitCentral_, pdFecIni, pdFecFin, psMoneda, psAgeCod)
    Do While Not rsQryEnc.EOF
        'pdFechaControl = Day(rsQryEnc!dEstadCTS)
        lnDiaMes = rsQryEnc!dEstadCTS - pdFecIni + 1
        'If Day(pdFechaControl) <> lnDiaMes Then
'          MsgBox "No se encontraron datos Estadisticos CTS del día " & pdFechaControl & "." & Chr(10) & "Por favor Consolide Estadísticas", vbInformation, "¡Aviso!"
        'Else
            If pnBitCentral_ = True Then
                If lnTipo = 1 Then
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(laSdoEnc(pnContAge, 3, lnDiaMes)) + nVal(rsQryEnc!nSaldoSinCmac) - nVal(rsQryEnc!MontoChq)
                    laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + nVal(rsQryEnc!nSaldoSinCmac) - nVal(rsQryEnc!MontoChq)
                End If
            Else
                If lnTipo = 1 Then
                    laSdoEnc(pnContAge, 3, lnDiaMes) = nVal(laSdoEnc(pnContAge, 3, lnDiaMes)) + nVal(rsQryEnc!nSaldo)
                    laTotEnc(pnContAge, 3) = Val(laTotEnc(pnContAge, 3)) + nVal(rsQryEnc!nSaldo)
                End If
            End If
            'rsQryEnc.MoveNext
        'End If
        rsQryEnc.MoveNext
        pdFechaControl = pdFechaControl + 1
    Loop
    RSClose rsQryEnc
End Sub

Private Function PrtSdoEnc(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    
    ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
    
    lsRTFEnc = ""
    lnLinPag = 65
    lnCarLin = UBound(laSdoEnc, 1) * 20   '  180
    If lnTipo = 1 Then
      lsTitRp1 = "E F E C T O S   D E   E N C A J E (CAPITAL)"
    Else
      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
    End If
    If nTipo = 2 Then
      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
    Else
      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    
    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "SEDE", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecIni, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & Space(11)
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
    
'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!Codigo, "00") & "    "
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    lsRTFEnc = lsRTFEnc & "   T O T A L    " & oImpresora.gPrnSaltoLinea
    xlHoja1.Cells(6, I + 1) = "TOTAL"
    
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lnLinPag = 8
    For I = LBound(laSdoEnc, 3) To UBound(laSdoEnc, 3)
      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoEnc(1, 1, I), gsFormatoFechaView) & " "
      xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, I), gsFormatoFechaView)
      
      nTotAge = 0
      For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
        lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoEnc(K, nTipo, I)), 15, 2) & " "
        xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoEnc(K, nTipo, I))
        nTotAge = nTotAge + nVal(laSdoEnc(K, nTipo, I))
      Next
      laProvis(1, I) = laSdoEnc(1, 1, I)
      laProvis(nTipo, I) = nTotAge
      laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
      lsRTFEnc = lsRTFEnc & PrnVal(nTotAge, 15, 2)
      xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    Next I
    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & String(11, " ")
    nTotAge = 0
    For K = LBound(laSdoEnc, 1) To UBound(laSdoEnc, 1)
        lsRTFEnc = lsRTFEnc & PrnVal(laTotEnc(K, nTipo), 15, 2) & " "
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotEnc(K, nTipo)
        nTotAge = nTotAge + laTotEnc(K, nTipo)
    Next
    
    lsRTFEnc = lsRTFEnc & PrnVal(nTotAge, 15, 2) & oImpresora.gPrnSaltoLinea
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
    PrtSdoEnc = lsRTFEnc
End Function
Private Function PrtSdoEncInt(nTipo As Integer, psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    
    ExcelAddHoja "Planilla_" & nTipo, xlLibro, xlHoja1
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
    lsRTFEnc = ""
    lnLinPag = 65
    lnCarLin = UBound(laSdoEncInt, 1) * 20   '  180
    If lnTipo = 1 Then
      lsTitRp1 = "E F E C T O S   D E   E N C A J E (INTERES)"
    Else
      lsTitRp1 = "P O S I C I O N   D E   L I Q U I D E Z"
    End If
    If nTipo = 2 Then
      lsTitRp1 = "A H O R R O   C O R R I E N T E   P A R A   " & lsTitRp1
    Else
      lsTitRp1 = "P L A Z O   F I J O   P A R A   " & lsTitRp1
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecIni, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea
               
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
               
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & Space(11)

'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!Codigo, "00") & "    "
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    lsRTFEnc = lsRTFEnc & "   T O T A L    " & oImpresora.gPrnSaltoLinea
    xlHoja1.Cells(6, I + 1) = "TOTAL"
    
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lnLinPag = 8
    For I = LBound(laSdoEncInt, 3) To UBound(laSdoEncInt, 3)
      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoEncInt(1, 1, I), gsFormatoFechaView) & " "
      xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoEnc(1, 1, I), gsFormatoFechaView)
      
      nTotAge = 0
      For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
        lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoEncInt(K, 2, I)), 15, 2) & "  "
        xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoEncInt(K, 2, I))
        
        nTotAge = nTotAge + nVal(laSdoEncInt(K, 2, I))
      Next
      'adiciona la provision para el total en la planilla provisional de encaje
      laProvis(nTipo, I) = nTotAge
      laTotPro(nTipo) = laTotPro(nTipo) + nTotAge
      lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2)
      xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
      
    Next I
    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & String(11, " ")
    nTotAge = 0
    For K = LBound(laSdoEncInt, 1) To UBound(laSdoEncInt, 1)
        lsRTFEnc = lsRTFEnc & PrnVal(Str(laTotEncInt(K, 1)), 15, 2) & "  "
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotEncInt(K, 1)
        
        nTotAge = nTotAge + laTotEncInt(K, 1)
    Next
    lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2) & oImpresora.gPrnSaltoLinea
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
    PrtSdoEncInt = lsRTFEnc
End Function

Private Function PrtSdoBillete(psMoneda As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnLinPag As Integer
    Dim lsCodMon As String
    Dim nTotAge As Currency
    
    ExcelAddHoja "Billetaje", xlLibro, xlHoja1
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = False
    xlHoja1.PageSetup.FitToPagesTall = 1
    xlHoja1.PageSetup.FitToPagesWide = 1
    xlHoja1.PageSetup.CenterVertically = True
    
    lsRTFEnc = ""
    lnLinPag = 65
    lnCarLin = (UBound(laSdoEnc, 1) * 20) + 18   '198
    If psMoneda = "1" Then
      lsTitRp1 = "C A J A    M O N E D A    N A C I O N A L"
    Else
      lsTitRp1 = "C A J A    M O N E D A    E X T R A N J E R A"
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    lsRTFEnc = lsRTFEnc & CabeRepo(CON & gsInstCmac, "", "CAJA GENERAL", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(ldFecFin, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnCntPag, 135) & oImpresora.gPrnSaltoLinea

    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & Space(11)
    
'Cabecera de Agencias
    Dim oDAge As New DActualizaDatosArea
    Set rs = oDAge.GetAgencias(, False, True)
    Set oDAge = Nothing
    I = LBound(laSdoEnc, 1)
    Do While Not rs.EOF
        lsRTFEnc = lsRTFEnc & "   Agencia " & Format(rs!Codigo, "00") & "    "
        xlHoja1.Cells(6, I + 1) = "Agencia " & Format(rs!Codigo, "00")
        I = I + 1
        rs.MoveNext
    Loop
    RSClose rs
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 10) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 10)).Font.Bold = True
    
    
    lsRTFEnc = lsRTFEnc & "  CAJA GENERAL          T O T A L "
    xlHoja1.Cells(6, I + 1) = "CAJA GENERAL"
    xlHoja1.Cells(6, I + 2) = "TOTAL"
        
   
    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lnLinPag = 8
    For I = LBound(laSdoBil, 2) To UBound(laSdoBil, 2)
      lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & Format(laSdoBil(UBound(laSdoBil, 1), I), gsFormatoFechaView) & " "
      xlHoja1.Cells(I + lnLinPag, 1) = "'" & Format(laSdoBil(UBound(laSdoBil, 1), I), gsFormatoFechaView)
      
      nTotAge = 0
      For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
          lsRTFEnc = lsRTFEnc & PrnVal(nVal(laSdoBil(K, I)), 15, 2) & "  "
          xlHoja1.Cells(I + lnLinPag, K + 1) = nVal(laSdoBil(K, I))
          
          nTotAge = nTotAge + nVal(laSdoBil(K, I))
      Next
      If psMoneda = "2" Then
         laProvis(6, I) = nTotAge
         laTotPro(6) = laTotPro(6) + nTotAge
      End If
      lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2)
      xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    Next I
    lsRTFEnc = lsRTFEnc & oImpresora.gPrnSaltoLinea & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFEnc = lsRTFEnc & String(11, " ")
    xlHoja1.Cells(I + lnLinPag, 1) = "Totales"
    
    nTotAge = 0
    For K = LBound(laSdoBil, 1) To UBound(laSdoBil, 1) - 1
        lsRTFEnc = lsRTFEnc & PrnVal(Str(laTotBil(K)), 15, 2) & "  "
        xlHoja1.Cells(I + lnLinPag, K + 1) = laTotBil(K)
        nTotAge = nTotAge + laTotBil(K)
    Next
    lsRTFEnc = lsRTFEnc & PrnVal(Str(nTotAge), 15, 2) & oImpresora.gPrnSaltoLinea
    xlHoja1.Cells(I + lnLinPag, K + 1) = nTotAge
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(I + lnLinPag, K + 1)).NumberFormat = "#,##0.00#"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, K + 1)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag, 1), xlHoja1.Cells(I + lnLinPag, K + 1)).BorderAround xlContinuous, xlMedium
    
    Dim oEst As New NEstadisticas
    If Day(ldFecFin + 1) = 1 Then  'Es ultimo dia del mes
       oEst.EliminaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda
       oEst.InsertaEstadAnexos ldFecFin, "CAJA_PROMEDIO", psMoneda, Round(nTotAge / Day(ldFecFin), 2)
    End If
    Set oEst = Nothing
    
    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=") & oImpresora.gPrnCondensadaOFF
    PrtSdoBillete = lsRTFEnc
End Function

Private Function PrtPlanillaProvision(psMoneda As String, lsFecha As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet) As String
    Dim lnCarLin As Integer, I As Integer, K As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lsCodMon As String
    
    'Variables Calculo de Encaje
    Dim R             As ADODB.Recordset
    Dim MatTemp()     As String
    Dim MatTemp2()    As String
    Dim MatTemp3()    As String
    Dim Total1        As Double
    Dim Total2        As Double
    Dim lnTotalEncaje As Double
    Dim pdFecIni As Date
    Dim oPara As New NEncajeBCR
    Dim lnLinPag As Long

    pdFecIni = Format("01" & Mid(lsFecha, 3), "MM/dd/YYYY")
    
    'Variables Calculo Encaje M.E.
'    Dim lnSEncProm As Double, lnSEncMarg As Double
'    Dim lnTasaBase As Double, lnTasaEncMarg As Double
'    Dim lnEncMarg As Double, lnEncCalc As Double
'    Dim lnTObligEnc As Double, lnEncExig As Double
    
'    Dim lnTasaEncLegMN As Double
'    Dim lnEncLeg As Double
    
    ExcelAddHoja "Planilla", xlLibro, xlHoja1
    xlHoja1.PageSetup.Zoom = 80
    
    lsRTFEnc = ""
    lnLinPag = 65
    lnCarLin = 140
    
'    Dim oBala As NBalanceCont
'    Set oBala = New NBalanceCont
'    lnTasaEncLegMN = oBala.GetUtilidadAcumulada("P", -1, Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
    
    
    
    If psMoneda = "1" Then
      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA NACIONAL"
    Else
      lsTitRp1 = "PLANILLA PROVISIONAL DE ENCAJE EN MONEDA EXTRANJERA"
    End If
    lsTitRp2 = "( DEL " & ldFecIni & " AL " & ldFecFin & " )"
    
    xlHoja1.Cells(1, 1) = gsInstCmac
    xlHoja1.Cells(1, 8) = ldFecFin
    xlHoja1.Cells(2, 1) = "CAJA GENERAL" & "-" & IIf(psMoneda = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(3, 2) = lsTitRp1
    xlHoja1.Cells(4, 2) = lsTitRp2
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 8)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 8)).Font.Bold = True

    xlHoja1.Cells(6, 1) = "FECHA"
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 1)).MergeCells = True
    
    xlHoja1.Cells(6, 2) = "DEPOSITOS Y OBLIGACIONES  SUJETOS DE ENCAJE": xlHoja1.Cells(7, 2) = "AHORRO": xlHoja1.Cells(7, 3) = "PLAZO FIJO": xlHoja1.Cells(7, 4) = "INTERES": xlHoja1.Cells(7, 5) = "TOTAL"
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).MergeCells = True
    
    xlHoja1.Cells(6, 6) = "EFECTIVO Y DEPOSITOS DE ENCAJE":  xlHoja1.Cells(7, 6) = "CAJA": xlHoja1.Cells(7, 7) = "DEPOSITOS BCR":     xlHoja1.Cells(7, 8) = "TOTAL"
    xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    lnLinPag = 9
    
    For I = LBound(laProvis, 2) To UBound(laProvis, 2)
      xlHoja1.Cells(I + lnLinPag - 1, 1) = "'" & Format(laProvis(1, I), gsFormatoFechaView)
      For K = 2 To 8
          xlHoja1.Cells(I + lnLinPag - 1, K) = nVal(laProvis(K, I))
      Next
    Next I
    xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(I + lnLinPag - 2, 8)).BorderAround xlContinuous, xlMedium
    
    For K = 2 To 8
        xlHoja1.Cells(I + lnLinPag - 1, K) = laTotPro(K)
    Next
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag - 1, 1), xlHoja1.Cells(I + lnLinPag - 1, 8)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(I + lnLinPag - 1, 8)).NumberFormat = "#,##0.00"
    
    'Imprimimos Cómputo del Encaje
    xlHoja1.Cells(I + lnLinPag + 1, 2) = "DIAS"
    xlHoja1.Cells(I + lnLinPag + 1, 3) = "ENCAJE CMAC"
    xlHoja1.Cells(I + lnLinPag + 1, 4) = "ENCAJE LEGAL"
    xlHoja1.Cells(I + lnLinPag + 1, 5) = "EXCEDENTE/DEFICIT"
    
'    If psMoneda = "1" Then
'      lnEncLeg = Round(laTotPro(5) * lnTasaEncLegMN / 100, 2) & "   "
'    Else
'      lnTObligEnc = oBala.GetUtilidadAcumulada("T", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'      lnSEncProm = Round(lnTObligEnc / 31, 2)
'      lnSEncMarg = Round((laTotPro(5) / nDias) - lnSEncProm, 2)
'      lnEncExig = oBala.GetUtilidadAcumulada("E", CInt(psMoneda), Month(CDate(lsFecha)), Year(CDate(ldFecFin)), True)
'      lnTasaBase = Round(lnEncExig / lnTObligEnc, 6)
'      If CDate(lsFecha) >= CDate("01/09/2000") And CDate(lsFecha) <= CDate("30/11/2000") Then
'            lnTasaBase = lnTasaBase - 0.03
'      End If
'      lnTasaEncMarg = 0.2
'      If lnSEncMarg > 0 Then
'         lnEncMarg = Round(lnSEncMarg * lnTasaEncMarg, 2)
'      Else
'         lnEncMarg = Round((lnSEncProm + lnSEncMarg) * lnTasaBase, 2)
'      End If
'      lnEncCalc = Round(lnSEncProm * lnTasaBase, 2)
'      lnEncLeg = Round((lnEncCalc + lnEncMarg) * nDias, 2)
'
'
'    End If
    
'    For i = LBound(laProvis, 2) To UBound(laProvis, 2)
'        MsgBox nVal(laProvis(8, i))
'    Next
    
    If psMoneda = "2" Then
        Set R = oPara.GetParametroEncaje(pdFecIni)
        ReDim MatTemp(UBound(laProvis, 2))
        ReDim MatTemp2(UBound(laProvis, 2))
        Total1 = 0
        Total2 = 0
        R.Find "nCodigo=2"
      
        'Hallo Exceso diario EncajeActual - Base
        For I = 1 To UBound(laProvis, 2)
            MatTemp(I) = nVal(laProvis(5, I)) - R!nValor
        Next I
      
        Total1 = R!nValor
      
        R.Find "nCodigo=3"
        'Hallo El Basico : Base * TasaBase (21.6478)%
        For I = 1 To UBound(laProvis, 2)
            MatTemp2(I) = Total1 * R!nValor
        Next I
        Total1 = 0#
      
        R.Find "nCodigo=4"
        
        'Encaje Marginal: Marginal por el Exeso * 20%
        For I = 1 To UBound(MatTemp)
            MatTemp(I) = CDbl(MatTemp(I)) * R!nValor
        Next I
      
        'Total Encaje : En Basico mas Marginal
        For I = 1 To UBound(MatTemp)
            MatTemp2(I) = CDbl(MatTemp(I)) + CDbl(MatTemp2(I))
        Next I
      
        For I = 1 To UBound(MatTemp)
            Total1 = Total1 + CDbl(MatTemp2(I))
        Next I
        Total1 = CDbl(Format(Total1, "#0.00"))
          
        R.Close
        Set R = Nothing
    Else
        Set R = oPara.GetParametroEncaje(pdFecIni)
        R.Find "nCodigo=1"
        Total2 = R!nValor
        R.Close
        Set R = Nothing
    End If
    
    If psMoneda = "1" Then
       lnTotalEncaje = CCur(laTotPro(8)) * Total2
   Else
       lnTotalEncaje = Total1
   End If
    
    'Dias
    xlHoja1.Cells(I + lnLinPag + 2, 2) = nDias
    'Encaje CMAC
    xlHoja1.Cells(I + lnLinPag + 2, 3) = laTotPro(8)
    'Encaje Legal
    'xlHoja1.Cells(i + lnLinPag + 2, 4) = Round(lnEncLeg, 2)
    xlHoja1.Cells(I + lnLinPag + 2, 4) = Round(lnTotalEncaje, 2)
    'Excedente
    'xlHoja1.Cells(i + lnLinPag + 2, 5) = Round(laTotPro(8) - lnEncLeg, 2)
    xlHoja1.Cells(I + lnLinPag + 2, 5) = Round(laTotPro(8) - lnTotalEncaje, 2)
    
    
    
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 2, 2), xlHoja1.Cells(I + lnLinPag + 2, 5)).NumberFormat = "#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 3), xlHoja1.Cells(I + lnLinPag + 2, 4)).ColumnWidth = 14
    xlHoja1.Range(xlHoja1.Cells(I + lnLinPag + 1, 5), xlHoja1.Cells(I + lnLinPag + 2, 5)).ColumnWidth = 18
    
    Dim oEst As New NEstadisticas
    oEst.EliminaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda
    oEst.InsertaEstadAnexos CDate(lsFecha), "ENCAJE_ACUMULA", psMoneda, Round(laTotPro(8) - lnTotalEncaje, 2)
    Set oEst = Nothing
    
    PrtPlanillaProvision = lsRTFEnc
    'Set oBala = Nothing
End Function

Private Sub CalculaTotProvis()
Dim K As Integer
Dim nTot1 As Currency
Dim nTot2 As Currency
laTotPro(5) = 0: laTotPro(8) = 0
For K = 1 To UBound(laProvis, 2)
   laProvis(5, K) = Val(laProvis(2, K)) + Val(laProvis(3, K) + Val(laProvis(4, K)))
   laProvis(8, K) = Val(laProvis(6, K)) + Val(laProvis(7, K))
   laTotPro(5) = laTotPro(5) + Val(laProvis(5, K))
   laTotPro(8) = laTotPro(8) + Val(laProvis(8, K))
Next
End Sub


''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''

Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
Dim pbActivaHoja   As Boolean

On Error Resume Next
pbActivaHoja = False
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
        If Not pbActivaHoja Then
            SendKeys "{ENTER}"
            xlHoja1.Delete
        Else
            xlHoja1.Activate
            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
            lbExisteHoja = True
        End If
       Exit For
    End If
Next
If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If Err Then
        Err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
End If
End Sub

Private Function ImprimeAdeudaCalendarioCabecera(lnNumPaginas As Integer, lnCapital As Currency, ldFechaApertura As Date, lnPlazoCuotas As Integer, _
       lnCuotas As Integer, lnGracia As Integer, lnInteres As Currency, lnPeriodo As Long, psCodAdeudado As String, lsEntidad As String, ByRef pnLin As Integer) As String
Dim lsTexto As String
lsTexto = ""
If pnLin > 60 Then
    lsTexto = lsTexto & PrnSet("B+") + Cabecera("SIMULACION DE CALENDARIOS DE PAGOS DE ADEUDADOS", lnNumPaginas, "", 80, , , lsEmprLogo) + PrnSet("B-") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & ImpreFormat("Adeudado N° : " & Trim(psCodAdeudado) & "-" & Trim(lsEntidad), 80) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & ImpreFormat("Capital     : " & ImpreFormat(Format(lnCapital, "#,#0.00"), 15, , True), 50) & "Fecha de Desembolso: " & ImpreFormat(Format(ldFechaApertura, "dd/mm/yyyy"), 12) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & ImpreFormat("Plazo       : " & ImpreFormat(Val(lnPlazoCuotas), 4, 0) & " Dias ", 50) & "Interes            :  " & ImpreFormat(lnInteres, 3, 2) & " % " & IIf(lnPeriodo > 30, "Anual", "Mensual") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & ImpreFormat("Cuotas      : " & ImpreFormat(lnCuotas, 4, 0), 50) & "Vencimiento Inicial: " & ImpreFormat(lnGracia, 4, 0) & " Dias " & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & PrnSet("B+") & String(105, "-") & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & ImpreFormat("Cuota", 8) & ImpreFormat("Fecha Pago", 15) & ImpreFormat("Capital", 12) & ImpreFormat("Interes", 11) & ImpreFormat("Comisión", 11) & ImpreFormat("TOTAL ", 13) & ImpreFormat("Saldo", 12) & oImpresora.gPrnSaltoLinea
    lsTexto = lsTexto & String(105, "-") & PrnSet("B-") & oImpresora.gPrnSaltoLinea
    pnLin = 13
End If
    ImprimeAdeudaCalendarioCabecera = lsTexto
End Function

Public Function ImprimeAdeudaCalendario(prs As Recordset, pnCapital As Currency, pdFecha As Date, pnPlazoCuotas As Integer, _
                        pnCuotas As Integer, pnGracia As Integer, pnInteres As Currency, pnPeriodo As Long, psCodAdeudado As String, psDescEntidad As String, pnTotCapital As Currency, pnTotInteres As Currency, pnTotGeneral As Currency)
    Dim nLineas As Integer
    Dim lnNumPaginas As Integer
    Dim lsImpre As String
    lsImpre = ""
    nLineas = 66
    lnNumPaginas = 0
    If Not prs Is Nothing Then
        If prs.State = adStateOpen Then
            Do While Not prs.EOF
                If nLineas > 62 Then
                    Linea lsImpre, ImprimeAdeudaCalendarioCabecera(lnNumPaginas, pnCapital, Format(pdFecha, gsFormatoFecha), pnPlazoCuotas, _
                        pnCuotas, pnGracia, pnInteres, pnPeriodo, psCodAdeudado, psDescEntidad, nLineas), 0
                End If
                Linea lsImpre, ImpreFormat(Val(prs.Fields(1)), 6, 0) & "  " & ImpreFormat(prs.Fields(0), 14, 0) & PrnSet("B+") & ImpreFormat(nVal(prs.Fields(2)), 12, , True) & ImpreFormat(prs.Fields(3), 10, , True) & PrnSet("B-") & ImpreFormat(nVal(prs.Fields(4)), 10, , True) & ImpreFormat(nVal(prs.Fields(5)), 12, , True) & ImpreFormat(nVal(prs.Fields(6)), 12, , True)
                nLineas = nLineas + 1
                prs.MoveNext
            Loop
        End If
    End If
    Linea lsImpre, String(90, "=") & PrnSet("B-")
    Linea lsImpre, PrnSet("B+") & ImpreFormat("Total Capital : " & Format(pnTotCapital, "#,#0.00"), 30) & ImpreFormat("Total Interes : " & Format(pnTotInteres, "#,#0.00"), 30) & ImpreFormat(" TOTAL  : " & Format(pnTotGeneral, "#,#0.00"), 30) & PrnSet("B-")
  ImprimeAdeudaCalendario = lsImpre
End Function



