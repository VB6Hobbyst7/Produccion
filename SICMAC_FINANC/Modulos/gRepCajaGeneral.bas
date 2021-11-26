Attribute VB_Name = "gRepCajaGeneral"
Option Base 1
Option Explicit
Dim oBarra As clsProgressBar

Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Sub ImprimeCartasFianza(pdFechaDel As Date, pdFechaAl As Date, pbIngreso As Boolean)
Dim oCaja   As New nCajaGenImprimir
Dim lsImpre As String
Dim lsMensaje As String

lsImpre = oCaja.ImprimeCartasFianza(pdFechaDel, pdFechaAl, pbIngreso, gsOpeCod, gdFecSis, lsMensaje)
If lsMensaje <> "" Then
    MsgBox lsMensaje, vbInformation, "Aviso"
    Exit Sub
End If

EnviaPrevio lsImpre, "REPORTE DE CARTAS FIANZA", gnLinPage, False
Set oCaja = Nothing

End Sub

Public Sub ResumenFlujoDiario(pdFecha As Date)
Dim oCaja   As New nCajaGenImprimir
Dim lsImpre As String
On Error GoTo ResumenFlujoDiarioErr
    lsImpre = oCaja.ImprimeResumenFlujoDiario(pdFecha, gsOpeCod)
    Set oCaja = Nothing
    EnviaPrevio lsImpre, "RESUMEN DIARIO DEL FLUJO DE CAJA ", gnLinPage, False
Exit Sub
ResumenFlujoDiarioErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Public Sub DetalleFlujoDiario(pdFechaDel As Date, pdFechaAl As Date)
Dim oCaja   As New nCajaGenImprimir
Dim lsImpre As String
On Error GoTo DetalleFlujoDiarioErr
    lsImpre = oCaja.ImprimeDetalleFlujoDiario(pdFechaDel, pdFechaAl, gsOpeCod)
    Set oCaja = Nothing
    EnviaPrevio lsImpre, "REPORTE DETALLADO DE FLUJO DE CAJA GENERAL", gnLinPage, False
Exit Sub
DetalleFlujoDiarioErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Public Sub ResumenCheques(pdFecha As Date, psOpeCod As String, Optional pbDetalle As Boolean = False)
    Dim lFecha As String
    Dim sCtaCod As String
    Dim nTotalDep As Currency
    Dim lsDocNro As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim nTotal As Currency
    Dim nTotG  As Currency
    Dim sTexto As String
    Dim sSql   As String
    Dim lsAge  As String
    
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oCont As NContImprimir
    Set oCont = New NContImprimir
    
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Set rs = New ADODB.Recordset
    Set rs = oOpe.CargaOpeCta(psOpeCod)
    
    oCon.AbreConexion
    
    If rs.EOF Then
        MsgBox "No se asignó Cuenta de Caja a Operación", vbCritical, "Error"
        Exit Sub
    End If
    sCtaCod = rs!cCtaContCod
    RSClose rs

    lFecha = pdFecha
    nTotal = 0
    sTexto = ""
    If pbDetalle Then
'Operaciones de Cheques.. se cambio filtro por LEFT(cOpeCod,1) = 4
' IN ('" & gOpeCGOpeBancosRegChequesMN & "','" & gOpeCGOpeBancosRegChequesME & "','" & gOpeCGOpeBancosRecibeChqAgMn & "','" & gOpeCGOpeBancosRecibeChqAgMe & "','" & gOpeCGOpeBancosDepChequesMN & "','" & gOpeCGOpeBancosDepChequesME & "','401153', '401)

        sSql = "SELECT aa.cAreaCod+ISNULL(AG.cAgeCod,'') cAgeCod, " _
             & "   ISNULL(AG.cAgeDescripcion,AA.cAreaDescripcion) cAgeDescripcion, " _
             & "   a.nTpoDoc, a.cNroDoc, a.cIFCta, md.dDocFecha, dValorizacion, p.cPersCod, p.cPersNombre, " _
             & "   SUM(a.nMonto) as nImporte, " _
             & "   SUM(Case a.nMoneda When 1 Then a.nMonto Else Round(a.nMonto * Isnull(MTC.nMovTpoCambio,0),2) End) as nImporteTC " _
             & "FROM DocRec a JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc " _
             & "       JOIN MOV M ON M.nMovNro = MD.nMovNro JOIN Persona P ON P.cPerscod = a.cPerscod " _
             & "       JOIN DocRecEst drc ON drc.cNroDoc = a.cNroDoc and drc.nTpoDoc = a.nTpoDoc and drc.cPersCod = a.cPersCod and drc.cIFCta = a.cIFCta and drc.cMovNro = m.cMovNro " _
             & "  Left Join MOVTPOCambio MTC ON M.nMovNro = MTC.nMovNro " _
             & "  LEFT JOIN Agencias AG ON a.cAgeCod = AG.cAgeCod " _
             & "  LEFT JOIN Areas AA ON a.cAreaCod = AA.cAreaCod " _
             & " WHERE drc.nEstado in (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & ") and a.nMoneda = " & Mid(psOpeCod, 3, 1) & " and M.nMovFlag not in (1) and datediff(dd, MD.dDocFecha, '" & Format(pdFecha, gsFormatoFecha) & "') = 0" _
             & " and (a.cDepIF = '0'" _
             & " or not exists" _
             & "(select mr.nMovNro from movref mr where mr.nMovNroRef = M.nMovNro" _
             & "        and substring(M.cMovNro,1,8) < '" & Format(pdFecha, gsFormatoMovFecha) & "') )" _
             & " GROUP BY aa.cAreaCod, AG.cAgeCod, AG.cAgeDescripcion, AA.cAreaDescripcion, " _
             & "  a.nTpoDoc , a.cNroDoc, a.cIFCta, md.dDocFecha, dValorizacion, p.cPersCod, p.cPersNombre " _
             & " ORDER BY aa.cAreaCod, AG.cAgeCod, a.cNroDoc"
    Else
        sSql = " SELECT aa.cAreaCod+ISNULL(AG.cAgeCod,'') cAgeCod, ISNULL(AG.cAgeDescripcion,AA.cAreaDescripcion) cAgeDescripcion, SUM(a.nMonto) as nImporte , SUM(Case a.nMoneda When 1 Then a.nMonto Else Round(a.nMonto * Isnull(MTC.nMovTpoCambio,0),2) End) as nImporteTC" _
             & " FROM DocRec a" _
             & " JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc" _
             & " JOIN MOV M ON M.nMovNro = MD.nMovNro" _
             & " JOIN DocRecEst drc ON drc.cNroDoc = a.cNroDoc and drc.nTpoDoc = a.nTpoDoc and drc.cPersCod = a.cPersCod and drc.cIFCta = a.cIFCta and drc.cMovNro = m.cMovNro " _
             & " Left Join MOVTPOCambio MTC ON M.nMovNro = MTC.nMovNro" _
             & " LEFT JOIN Agencias AG ON a.cAgeCod = AG.cAgeCod LEFT JOIN Areas AA ON a.cAreaCod = AA.cAreaCod " _
             & " WHERE drc.nEstado in (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & ") and a.nMoneda = " & Mid(psOpeCod, 3, 1) & " and M.nMovFlag not in (1) and datediff(dd, MD.dDocFecha, '" & Format(pdFecha, gsFormatoFecha) & "') = 0" _
             & " and (a.cDepIF = '0'" _
             & " or not exists" _
             & "(select mr.nMovNro from movref mr where mr.nMovNroRef = M.nMovNro" _
             & "        and substring(M.cMovNro,1,8) < '" & Format(pdFecha, gsFormatoMovFecha) & "') )" _
             & " GROUP BY aa.cAreaCod, AG.cAgeCod, AG.cAgeDescripcion, AA.cAreaDescripcion"
    End If
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        MsgBox "No se registraron Cheques a la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    If Mid(psOpeCod, 3, 1) = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    gsMovNro = Format(lFecha & " " & Time, gsFormatoMovFechaHora)
    nTotal = 0
    
    If pbDetalle Then
        sTexto = ImpreCabAsiento(80, lFecha, gsNomCmac, psOpeCod, "", "REPORTE DETALLADO DIARIO DE CHEQUES RECIBIDOS", True) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "=") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & " Instituc. Financ.                         Nro.Doc   Fec.Reg.    Fec.Valoriz.    Monto" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "-") & oImpresora.gPrnSaltoLinea
        lsAge = ""
        Do While Not rs.EOF
            If lsAge <> rs!cAgeCod Then
                If nTotal <> 0 Then
                    sTexto = sTexto & space(51) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
                    sTexto = sTexto & oImpresora.gPrnSaltoLinea
                    nTotal = 0
                End If
                sTexto = sTexto & " " & "AGENCIA : " & Justifica(rs!cAgeCod, 5) & "  " & rs!cAgeDescripcion & oImpresora.gPrnSaltoLinea
                lsAge = rs!cAgeCod
            End If
            sTexto = sTexto & " " & Justifica(rs!cPersNombre, 35) & " " & Justifica(rs!cNroDoc, 16) & " " & rs!dDocFecha & " " & rs!dValorizacion & " " & PrnVal(rs!nImporte, 14, 2) & oImpresora.gPrnSaltoLinea
            nTotal = nTotal + rs!nImporte
            nTotG = nTotG + rs!nImporte
            rs.MoveNext
        Loop
        sTexto = sTexto & space(51) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
        sTexto = sTexto & oImpresora.gPrnSaltoLinea
        nTotal = 0
    Else
        sTexto = ImpreCabAsiento(80, lFecha, gsNomCmac, psOpeCod, "", "RESUMEN DIARIO DE CHEQUES RECIBIDOS", True) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(78, "=") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & " Código      Descripción                                             Monto" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(78, "-") & oImpresora.gPrnSaltoLinea
        Do While Not rs.EOF
            sTexto = sTexto & " " & Justifica(rs!cAgeCod, 5) & "  " & Mid(rs!cAgeDescripcion & space(43), 1, 43) & space(12) & Right(space(14) & Format(rs!nImporte, gsFormatoNumeroView), 14) & oImpresora.gPrnSaltoLinea
            nTotal = nTotal + rs!nImporte
            nTotG = nTotG + rs!nImporte
            rs.MoveNext
        Loop
    End If
    RSClose rs
    sTexto = sTexto & String(78, "-") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & BON & Mid(" TOTAL CHEQUES RECIBIDOS EN EL DIA" & space(60), 1, 59) & gsSimbolo & " " & Right(space(14) & Format(nTotG, gsFormatoNumeroView), 14) & BOFF & oImpresora.gPrnSaltoLinea

    sSql = " SELECT ISNUll(SUM(ISNUll(a.nMonto,0)),0) as nImporte , ISNUll(SUM(ISNUll(Case a.nMoneda When 1 Then a.nMonto Else Round(a.nMonto * Isnull(MTC.nMovTpoCambio,0),2) End,0)),0) as nImporteTC" _
         & " FROM   DocRec a" _
         & " JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc" _
         & " JOIN MOV M ON M.nMovNro = MD.nMovNro JOIN DocRecEst drc ON drc.cNroDoc = a.cNroDoc and drc.nTpoDoc = a.nTpoDoc and drc.cPersCod = a.cPersCod and drc.cMovNro = m.cMovNro " _
         & "      and drc.cIFCta = a.cIFCta " _
         & " " _
         & " Left Join MOVTPOCambio MTC ON M.nMovNro = MTC.nMovNro" _
         & " JOIN Agencias AG ON Substring(M.cMovNro,18,2) = AG.cAgeCod" _
         & " WHERE drc.nEstado in (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & ") and  nMoneda = " & Mid(psOpeCod, 3, 1) & " And MD.dDocFecha < '" & Format(pdFecha, gsFormatoFecha) & "' and m.nMovFlag <> 1 and " _
         & " (a.cDepIF = '0' or EXISTS" _
         & " (SELECT mr.nMovNro FROM MovRef MR" _
         & " Inner Join Mov MOVR On MOVR.nMovNro = MR.nMovNro" _
         & " WHERE movr.nMovFlag <> 1 and MR.nMovNroRef = M.nMovNro and RTRIM(ISNULL(MR.cAgeCodRef,'')) = '' and substring(MOVR.cmovnro,1,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "'))"

    Set rs = oCon.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        sTexto = sTexto & BON & Mid(" TOTAL CHEQUES NO DEPOSITADOS DEL DIA ANTERIOR" & space(60), 1, 59) & gsSimbolo & " " & PrnVal(rs!nImporte, 14, 2) & BOFF & oImpresora.gPrnSaltoLinea
        nTotG = nTotG + rs!nImporte
    End If
    
    sSql = "SELECT distinct c.nMonto as nImporte,  ISNUll(Case c.nMoneda When 1 Then c.nMonto Else Round(c.nMonto * Isnull(MTC.nMovTpoCambio,0),2) End,0) as nImporteTC, cDocNro " _
         & "FROM Mov m JOIN MovCta a ON a.nMovNro = m.nMovNro LEFT JOIN MovTpoCambio mtc ON mtc.nMovNro = a.nMovNro " _
         & "      JOIN MovRef mr ON mr.nMovNro = m.nMovNro " _
         & "      JOIN MovDoc md ON md.nMovNro = mr.nMovNroRef " _
         & "      JOIN DocRec c ON c.cNroDoc = md.cDocNro " _
         & "WHERE cCtaContCod LIKE '" & sCtaCod & "%' and " _
         & "      substring(m.cmovnro,1,8) = '" & Format(lFecha, gsFormatoMovFecha) & "' and a.nMovImporte < 0 and " _
         & "      Datediff(day,md.dDocFecha,'" & Format(lFecha, gsFormatoFecha) & "')=0 "

    Set rs = oCon.CargaRecordSet(sSql)
    nTotalDep = 0
    If Not rs.EOF Then
        lsDocNro = ""
        Do While Not rs.EOF
            If lsDocNro <> Trim(rs!cDocNro) Then
                nTotalDep = nTotalDep + IIf(IsNull(rs!nImporte), 0, rs!nImporte)
            End If
            If Not rs.EOF Then
                lsDocNro = Trim(rs!cDocNro)
            End If
            rs.MoveNext
        Loop
    End If
    sTexto = sTexto & BON & Mid(" TOTAL CHEQUES DEPOSITADOS EN EL DIA" & space(60), 1, 59) & gsSimbolo & " " & Right(space(14) & Format(nTotalDep, gsFormatoNumeroView), 14) & BOFF & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & String(78, "=") & oImpresora.gPrnSaltoLinea
    If nTotG - nTotalDep >= 0 Then
        sTexto = sTexto & BON & Mid(" TOTAL CHEQUES EN CARTERA" & space(60), 1, 59) & gsSimbolo & " " & Right(space(14) & Format(nTotG - nTotalDep, gsFormatoNumeroView), 14) & BOFF & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(78, "=") & oImpresora.gPrnSaltoLinea
    End If

    sTexto = sTexto & ""
    oCon.CierraConexion
    oPrevio.Show sTexto, "Resumen de Cheques Recibidos", False, gnLinPage, gImpresora
End Sub
Public Sub ResumenChequesRem(pdFechaDel As Date, pdFechaAl As Date, psOpeCod As String)
    Dim lFecha As String
    Dim sCadena As String
    Dim nTotalDep As Currency
    Dim lsDocNro As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim nTotal As Currency
    Dim nTotG  As Currency
    Dim sTexto As String
    Dim sSql   As String
    Dim lsAge  As String
    
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oCont As NContImprimir
    Set oCont = New NContImprimir
    
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Set rs = New ADODB.Recordset
    
     If Mid(psOpeCod, 3, 1) = "1" Then
        gsSimbolo = "1"
    Else
        gsSimbolo = "2"
    End If
    
    sCadena = "42" & gsSimbolo & "210"
    
    oCon.AbreConexion
    

 lFecha = gdFecSis
    nTotal = 0
    sTexto = ""

sSql = "Select MOA.cAgeCod,AG.cAgeDescripcion as cAgencia,MD.nDocTpo,MD.cDocNro as cCheque,CTA.cCtaIfdesc,MD.dDocFecha, " _
    & " Pers.cPersCod,Pers.cPersNombre,sum(MOV.nMovImporte) as nMonto" _
    & " From MovDoc MD" _
    & " Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro " _
    & " Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod" _
    & " Inner Join CtaIF CTA on CTA.cPerscod=MOI.cPerscod and CTA.cIFTpo=MOI.cIFTpo and CTA.cCtaIfCod=MOI.cCtaIfCod" _
    & " Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro" _
    & " Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod" _
    & " Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro " _
    & " Inner Join Mov M on M.nMovNro=MD.nMovNro " _
    & " Where MD.nDocTpo='47' and M.nMovFlag=0 and M.copecod ='" & sCadena & "' and substring(M.cmovnro,1,8)  BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' AND '" & Format(pdFechaAl, "yyyymmdd") & "'  " _
    & " and Not Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)   " _
    & " GROUP BY MOA.cAgeCod,AG.cAgeDescripcion,MD.nDocTpo ,MD.cDocNro, CTA.cCtaIFdesc, " _
    & " md.dDocFecha , pers.cPersCod, pers.cPersNombre " _
    & " ORDER BY MOA.cAgeCod, MD.cDocNro " _
    
    
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        MsgBox "No se registraron Cheques a la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    gsMovNro = Format(lFecha & " " & Time, gsFormatoMovFechaHora)
    nTotal = 0
    
        sTexto = ImpreCabAsiento(80, lFecha, gsNomCmac, psOpeCod, "", "REPORTE DETALLADO DIARIO DE CHEQUES EMITIDOS", True) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & Centra(" DESDE " & pdFechaDel & " AL " & pdFechaAl) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "=") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & " Instituc. Financ.                   Nro.Doc         Fec.Reg.              Monto" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "-") & oImpresora.gPrnSaltoLinea
        lsAge = ""
        Do While Not rs.EOF
            If lsAge <> rs!cAgeCod Then
                If nTotal <> 0 Then
                    sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
                    sTexto = sTexto & oImpresora.gPrnSaltoLinea
                    nTotal = 0
                End If
                sTexto = sTexto & " " & "AGENCIA : " & Justifica(rs!cAgeCod, 5) & "  " & rs!cAgencia & oImpresora.gPrnSaltoLinea
                lsAge = rs!cAgeCod
            End If
            sTexto = sTexto & " " & Justifica(rs!cPersNombre, 35) & " " & Justifica(rs!cCheque, 16) & " " & rs!dDocFecha & "  " & PrnVal(rs!nMonto, 14, 2) & oImpresora.gPrnSaltoLinea
            nTotal = nTotal + rs!nMonto
            nTotG = nTotG + rs!nMonto
            rs.MoveNext
        Loop
        sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
        sTexto = sTexto & oImpresora.gPrnSaltoLinea
        nTotal = 0
    
    RSClose rs

    sTexto = sTexto & ""
    oCon.CierraConexion
    oPrevio.Show sTexto, "Resumen de Cheques Recibidos", False, gnLinPage, gImpresora
End Sub

Public Sub ResumenChequesAnul(pdFechaDel As Date, pdFechaAl As Date, psOpeCod As String)
    Dim lFecha As String
    Dim sCadena As String
    Dim nTotalDep As Currency
    Dim lsDocNro As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim nTotal As Currency
    Dim nTotG  As Currency
    Dim sTexto As String
    Dim sSql   As String
    Dim lsAge  As String
    
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oCont As NContImprimir
    Set oCont = New NContImprimir
    
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Set rs = New ADODB.Recordset
    
     If Mid(psOpeCod, 3, 1) = "1" Then
        gsSimbolo = "1"
    Else
        gsSimbolo = "2"
    End If
    
    sCadena = "42" & gsSimbolo & "210"
    
    oCon.AbreConexion
    

 lFecha = gdFecSis
    nTotal = 0
    sTexto = ""

sSql = "Select MOA.cAgeCod,AG.cAgeDescripcion as cAgencia,MD.nDocTpo,MD.cDocNro as cCheque,CTA.cCtaIfdesc,MD.dDocFecha, " _
    & " Pers.cPersCod,Pers.cPersNombre,sum(MOV.nMovImporte) as nMonto" _
    & " From MovDoc MD" _
    & " Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro " _
    & " Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod" _
    & " Inner Join CtaIF CTA on CTA.cPerscod=MOI.cPerscod and CTA.cIFTpo=MOI.cIFTpo and CTA.cCtaIfCod=MOI.cCtaIfCod" _
    & " Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro" _
    & " Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod" _
    & " Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro " _
    & " Inner Join Mov M on M.nMovNro=MD.nMovNro " _
    & " Where MD.nDocTpo='47' and M.nMovFlag=1 and M.copecod ='" & sCadena & "' and substring(M.cmovnro,1,8)  BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' AND '" & Format(pdFechaAl, "yyyymmdd") & "'  " _
    & " and Not Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)   " _
    & " GROUP BY MOA.cAgeCod,AG.cAgeDescripcion,MD.nDocTpo ,MD.cDocNro, CTA.cCtaIFdesc, " _
    & " md.dDocFecha , pers.cPersCod, pers.cPersNombre " _
    & " ORDER BY MOA.cAgeCod, MD.cDocNro " _
    
    
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        MsgBox "No se registraron Cheques a la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    gsMovNro = Format(lFecha & " " & Time, gsFormatoMovFechaHora)
    nTotal = 0
    
        sTexto = ImpreCabAsiento(80, lFecha, gsNomCmac, psOpeCod, "", "REPORTE DETALLADO DIARIO DE CHEQUES ANULADOS", True) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & Centra(" DESDE " & pdFechaDel & " AL " & pdFechaAl) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "=") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & " Instituc. Financ.                   Nro.Doc         Fec.Reg.              Monto" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "-") & oImpresora.gPrnSaltoLinea
        lsAge = ""
        Do While Not rs.EOF
            If lsAge <> rs!cAgeCod Then
                If nTotal <> 0 Then
                    sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
                    sTexto = sTexto & oImpresora.gPrnSaltoLinea
                    nTotal = 0
                End If
                sTexto = sTexto & " " & "AGENCIA : " & Justifica(rs!cAgeCod, 5) & "  " & rs!cAgencia & oImpresora.gPrnSaltoLinea
                lsAge = rs!cAgeCod
            End If
            sTexto = sTexto & " " & Justifica(rs!cPersNombre, 35) & " " & Justifica(rs!cCheque, 16) & " " & rs!dDocFecha & "  " & PrnVal(rs!nMonto, 14, 2) & oImpresora.gPrnSaltoLinea
            nTotal = nTotal + rs!nMonto
            nTotG = nTotG + rs!nMonto
            rs.MoveNext
        Loop
        sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
        sTexto = sTexto & oImpresora.gPrnSaltoLinea
        nTotal = 0
    
    RSClose rs

    sTexto = sTexto & ""
    oCon.CierraConexion
    oPrevio.Show sTexto, "Resumen de Cheques Recibidos", False, gnLinPage, gImpresora
End Sub

Public Sub ResumenChequesCob(pdFechaDel As Date, pdFechaAl As Date, psOpeCod As String)
    Dim lFecha As String
    Dim sCadena As String
    Dim nTotalDep As Currency
    Dim lsDocNro As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim nTotal As Currency
    Dim nTotG  As Currency
    Dim sTexto As String
    Dim sSql   As String
    Dim lsAge  As String
    
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oCont As NContImprimir
    Set oCont = New NContImprimir
    
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Set rs = New ADODB.Recordset
    
     If Mid(psOpeCod, 3, 1) = "1" Then
        gsSimbolo = "1"
    Else
        gsSimbolo = "2"
    End If
    
    sCadena = "42" & gsSimbolo & "210"
    
    oCon.AbreConexion
    

 lFecha = gdFecSis
    nTotal = 0
    sTexto = ""

sSql = "Select MOA.cAgeCod,AG.cAgeDescripcion as cAgencia,MD.nDocTpo,MD.cDocNro as cCheque,CTA.cCtaIfdesc,MD.dDocFecha, " _
    & " Pers.cPersCod,Pers.cPersNombre,sum(MOV.nMovImporte) as nMonto" _
    & " From MovDoc MD" _
    & " Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro " _
    & " Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod" _
    & " Inner Join CtaIF CTA on CTA.cPerscod=MOI.cPerscod and CTA.cIFTpo=MOI.cIFTpo and CTA.cCtaIfCod=MOI.cCtaIfCod" _
    & " Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro" _
    & " Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod" _
    & " Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro " _
    & " Inner Join Mov M on M.nMovNro=MD.nMovNro " _
    & " Where MD.nDocTpo='47' and M.nMovFlag=0 and M.copecod ='" & sCadena & "' and substring(M.cmovnro,1,8)  BETWEEN '" & Format(pdFechaDel, "yyyymmdd") & "' AND '" & Format(pdFechaAl, "yyyymmdd") & "'  " _
    & " and  Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)   " _
    & " GROUP BY MOA.cAgeCod,AG.cAgeDescripcion,MD.nDocTpo ,MD.cDocNro, CTA.cCtaIFdesc, " _
    & " md.dDocFecha , pers.cPersCod, pers.cPersNombre " _
    & " ORDER BY MOA.cAgeCod, MD.cDocNro " _
    
    
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        MsgBox "No se registraron Cheques a la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    gsMovNro = Format(lFecha & " " & Time, gsFormatoMovFechaHora)
    nTotal = 0
    
        sTexto = ImpreCabAsiento(80, lFecha, gsNomCmac, psOpeCod, "", "REPORTE DETALLADO DIARIO DE CHEQUES COBRADOS", True) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & Centra(" DESDE " & pdFechaDel & " AL " & pdFechaAl) & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "=") & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & " Instituc. Financ.                   Nro.Doc         Fec.Reg.              Monto" & oImpresora.gPrnSaltoLinea
        sTexto = sTexto & String(88, "-") & oImpresora.gPrnSaltoLinea
        lsAge = ""
        Do While Not rs.EOF
            If lsAge <> rs!cAgeCod Then
                If nTotal <> 0 Then
                    sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
                    sTexto = sTexto & oImpresora.gPrnSaltoLinea
                    nTotal = 0
                End If
                sTexto = sTexto & " " & "AGENCIA : " & Justifica(rs!cAgeCod, 5) & "  " & rs!cAgencia & oImpresora.gPrnSaltoLinea
                lsAge = rs!cAgeCod
            End If
            sTexto = sTexto & " " & Justifica(rs!cPersNombre, 35) & " " & Justifica(rs!cCheque, 16) & " " & rs!dDocFecha & "  " & PrnVal(rs!nMonto, 14, 2) & oImpresora.gPrnSaltoLinea
            nTotal = nTotal + rs!nMonto
            nTotG = nTotG + rs!nMonto
            rs.MoveNext
        Loop
        sTexto = sTexto & space(41) & "Total Cheques en Agencia " & PrnVal(nTotal, 14, 2)
        sTexto = sTexto & oImpresora.gPrnSaltoLinea
        nTotal = 0
    
    RSClose rs

    sTexto = sTexto & ""
    oCon.CierraConexion
    oPrevio.Show sTexto, "Resumen de Cheques Recibidos", False, gnLinPage, gImpresora
End Sub

Public Sub ResumenChqRecibidos(pdFecha As Date, psOpeCod As String)
    Dim lFecha As String
    Dim lsCodObj As String
    Dim nTotalDep As Currency
    Dim lsDocNro As String
    Dim lnTotal As Currency
    Dim sTexto  As String
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim oCon As DConecta
    Dim oOpe As DOperacion
    
    Set oCon = New DConecta
    Set oOpe = New DOperacion
    oCon.AbreConexion
    Set rs = oOpe.GetOpeObj(psOpeCod, 0)
    If rs.EOF Then
        MsgBox "No se asignó Objeto a Operación", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    lsCodObj = rs!Codigo
    RSClose rs

    lFecha = pdFecha
    lnTotal = 0
    sTexto = ""

    sSql = "    SELECT  CONVERT(CHAR(35),p.cPersNombre) as Banco, md.dDocFecha as Registro, " _
       & "         p.cPersCod, dr.cNroDoc, dr.dValorizacion, ce.cConsDescripcion as Estado, " _
       & "         dr.nMonto nMontoChq, dr.cIFCta cCtaBco, ISNULL(drc.cCtaCod,'') cCtaCod , cDepIF, RIGHT(m.cMovNro,4) as Usuario, " _
       & "         'CONDICION'= CASE  " _
       & "                      WHEN  cDepIF ='0' THEN 'NO DEPOSITADO' " _
       & "                      WHEN  cDepIF ='1' THEN 'DEPOSITADO' " _
       & "                      WHEN  cDepIF ='2' THEN 'DEVUELTO' End, " _
       & "         'PLAZA' = CASE WHEN bPlaza = '1' THEN 'Misma Plaza' ELSE 'Otra Plaza' END " _
       & " FROM   DocRec dr LEFT JOIN DocRecCapta drc ON dr.cPersCod = drc.cPersCod and dr.nTpoDoc = drc.nTpoDoc and dr.cNroDoc = drc.cNroDoc " _
       & "        JOIN MovDoc md on md.cDocNro = dr.cNroDoc and md.nDocTpo = dr.nTpoDoc  " _
       & "        JOIN Mov m ON m.nMovNro = md.nMovNro JOIN DocRecEst dre ON dre.cNroDoc = dr.cNroDoc and dre.nTpoDoc = dr.nTpoDoc and dre.cMovNro = m.cMovNro " _
       & "        JOIN Persona p ON p.cPersCod = dr.cPersCod " _
       & "        JOIN Constante ce ON ce.nConsValor = dr.nEstado and ce.nConsCod like '" & gChequeEstado & "' " _
       & " WHERE   dr.nMoneda = '" & Mid(psOpeCod, 3, 1) & "' and m.nMovFlag <> " & gMovFlagEliminado & " and m.cOpeCod in ('" & gOpeCGOpeBancosRegChequesMN & "','" & gOpeCGOpeBancosRegChequesME & "') " _
       & "         and dr.cAreaCod  = '" & Left(lsCodObj, 3) & "' and dr.cAgeCod = '" & Mid(lsCodObj, 4, 2) & "' and datediff(d,md.dDocFecha,'" & Format(lFecha, gsFormatoFecha) & "') = 0 " _
       & " ORDER BY Banco "

    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        rs.Close
        MsgBox "No se registraron Cheques a la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    lnTotal = 0
    sTexto = sTexto + PrnSet("C+") + PrnSet("B+") + CabeRepo(gsNomCmac, gsNomAge, "Caja General", "", lFecha, "LISTADO  DE CHEQUES RECIBIDOS EN CAJA GENERAL", "DATOS AL :" & lFecha & " EN " & IIf(Mid(psOpeCod, 3, 1) = "1", "MN", "ME"), "", "", 0, gnColPage) & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & String(145, "=") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & ImpreFormat("CUENTA", 14) & ImpreFormat("Nro Cheque", 14) & ImpreFormat("BANCO", 25) & ImpreFormat("CTA.BANCO", 14) & ImpreFormat("MONTO", 13) & ImpreFormat("FECHA REG.", 10) & ImpreFormat("FECHA VAL.", 10) & ImpreFormat("USUA", 4) & ImpreFormat("CONDICION", 10) & ImpreFormat("PLAZA", 12) & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & String(145, "=") + PrnSet("B-") & oImpresora.gPrnSaltoLinea
    Do While Not rs.EOF
        sTexto = sTexto & ImpreFormat(rs!cCtaCod, 14) & ImpreFormat(rs!cNroDoc, 14) & _
           ImpreFormat(rs!banco, 25) & ImpreFormat(rs!cCtaBco, 14) & ImpreFormat(rs!nMontoChq, 12, 2, True) & _
           ImpreFormat(Format(rs!Registro, "dd/mm/yyyy"), 10) & ImpreFormat(Format(rs!dValorizacion, "dd/mm/yyyy"), 10) & _
           ImpreFormat(rs!Usuario, 4) & ImpreFormat(rs!Condicion, 10) & ImpreFormat(rs!Plaza, 12) & oImpresora.gPrnSaltoLinea

        lnTotal = lnTotal + rs!nMontoChq
        rs.MoveNext
    Loop
    RSClose rs
    sTexto = sTexto & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + PrnSet("B+") + ImpreFormat("", 52) + ImpreFormat("TOTAL RECIBIDO : " & gsSimbolo, 17) & ImpreFormat(lnTotal, 12, 2, True) & oImpresora.gPrnSaltoLinea & PrnSet("C-")
    EnviaPrevio sTexto, "Resumen de Cheques Recibidos", gnLinPage, False
End Sub

Public Sub ReporteOrdenesPago(pdFecha As Date, pdFecha2 As Date, psOpeCod As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lsCadenaPrint As String
    Dim dFecha As Date
    Dim dFecha2 As Date
    Dim nLineas As Integer
    Dim lnPagina As Integer
    Dim lsCodCtaCont As String

    Dim oCon As DConecta
    Dim oOpe As DOperacion
    
    Set oCon = New DConecta
    Set oOpe = New DOperacion
    
    dFecha = pdFecha
    dFecha2 = pdFecha2

    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Set rs = New ADODB.Recordset
    Set rs = oOpe.CargaOpeCta(psOpeCod)
    
    If rs.EOF Then
        MsgBox "No se asignó Cuenta de Caja a Operación", vbCritical, "Error"
        Exit Sub
    End If
    lsCodCtaCont = rs!cCtaContCod
    RSClose rs

    oCon.AbreConexion
    sql = "SELECT  M.CMOVDESC, M.CMOVNRO, MD.nDOCTPO , MD.CDOCNRO, M.nMOVFLAG, " _
       & " MC.CCTACONTCOD, MC.NMOVIMPORTE nImporteMN , ISNULL(ME.NMOVMEIMPORTE,0) nImporteME, " _
       & " Estado= CASE WHEN M.nMOVFLAG = '1' THEN 'ANULADA' WHEN M.nMovFlag = '2' THEN 'EXTORNADA' WHEN m.nMovFlag = '3' THEN 'DE EXTORNO' " _
       & "             Else 'EMITIDA' End, ISNULL(ISNULL( ISNULL(mg.cPersCod,mar.cPersCod), ch.cPersCod ),Rend.cPersCod) cPersCod, " _
       & "         ISNULL(p.cPersNombre,'') cPersNombre " _
       & " FROM    MOV M JOIN MOVDOC MD ON M.nMOVNRO=MD.nMOVNRO " _
       & "         JOIN MOVCTA MC ON MC.nMOVNRO = M .nMOVNRO " _
       & "         LEFT JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
       & "         LEFT JOIN (Select mr.nMovNro, mar.cPersCod FROM MovRef mr " _
       & "                           JOIN MovArendir mar ON mar.nMovNro = mr.nMovNroRef " _
       & "                   ) mar ON mar.nMovNro = m.nMovNro " _
       & "         LEFT JOIN (SELECT mr.nMovNro, mar.cPersCod FROM movref mr join movref mr1 on mr1.nmovnro = mr.nmovnroref " _
       & "                         JOIN movarendir mar ON mar.nmovnro = mr1.nmovnroref ) Rend ON Rend.nMovNro = m.nMovNro " _
       & "         LEFT JOIN MovCajaChica mch ON mch.nMovnro = m.nMovNro and mch.cProcTpo = 8 " _
       & "         LEFT JOIN CajaChica ch ON ch.cAreaCod = mch.cAreaCod and ch.cAgeCod = mch.cAgeCod and ch.nProcNro = mch.nProcNro " _
       & "         LEFT JOIN Persona P on P.cPersCod = ISNULL(ISNULL( ISNULL(mg.cPersCod,mar.cPersCod), ch.cPersCod ),Rend.cPersCod) " _
       & "         LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM " _
       & " WHERE   MD.nDocTpo = '" & TpoDocOrdenPago & "' and m.nMovFlag <> '" & gMovFlagEliminado & "' and m.nMovEstado = '" & gMovEstContabMovContable & "' " _
       & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(dFecha, "yyyymmdd") & "' AND '" & Format(dFecha2, "yyyymmdd") & "' " _
       & "         AND SUBSTRING(MC.CCTACONTCOD,3,1)='" & Mid(psOpeCod, 3, 1) & "' AND MC.CCTACONTCOD LIKE '" & lsCodCtaCont & "%'  " _
       & " ORDER BY MD.CDOCNRO, M.CMOVNRO "

    lsCadenaPrint = ""
    nLineas = 0
    lnPagina = 0
    Set rs = oCon.CargaRecordSet(sql)
    If Not RSVacio(rs) Then
        nLineas = CabReporteOP(lsCadenaPrint, lnPagina, psOpeCod)
        Do While Not rs.EOF
            lsCadenaPrint = lsCadenaPrint & ImpreFormat(Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), 12) & _
               ImpreFormat(rs!cDocNro, 10) & ImpreFormat(rs!cPersNombre, 40) & ImpreFormat(rs!cMovDesc, 50) & _
               ImpreFormat(IIf(Mid(psOpeCod, 3, 1) = "2", Abs(rs!nImporteME), Abs(rs!nImporteMN)), 12, , True) & ImpreFormat(rs!estado, 8) & oImpresora.gPrnSaltoLinea
            nLineas = nLineas + 1
            If nLineas > 60 Then
'                lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
'                lnPagina = lnPagina + 1
                nLineas = CabReporteOP(lsCadenaPrint, lnPagina, psOpeCod)
            End If
            rs.MoveNext
            DoEvents
        Loop
        If lsCadenaPrint <> "" Then
            EnviaPrevio lsCadenaPrint, "Reporte de Ordenes de pagos Giradas", gnLinPage, True
        End If
    Else
        MsgBox "No se encontraron registros seleccionados", vbInformation, "Aviso"
    End If
    RSClose rs
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Private Function CabReporteOP(lsCadenaPrint As String, lnPagina As Integer, psOpeCod As String) As Integer
    lsCadenaPrint = lsCadenaPrint & CabeRepo(gsNomCmac, gsNomAge, "Area de Caja General", "", Format(gdFecSis, "dd/mm/yyyy"), "R E P O R T E   D E   O R D E N E S  D E  P A G O  G I R A D A S ", IIf(Mid(psOpeCod, 3, 1) = "1", "M O N E D A  N A C I O N A L", "M O N E D A  E X T R A N J E R A"), "", "", lnPagina, 150) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnBoldON & String(150, "-") & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat("Fecha Emision", 12) & ImpreFormat("Nro de O/P", 10) & ImpreFormat("PERSONA GIRADA", 40) & ImpreFormat("DESCRIPCION", 56) & ImpreFormat("IMPORTE", 8) & ImpreFormat("ESTADO", 8) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & String(150, "-") & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    CabReporteOP = 10
End Function

Public Sub ReporteArendirCuentaLibro(pdFecha As Date, pdFecha2 As Date, psOpeCod As String, Optional pbCajaChica As Boolean = False)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsFechaDesde  As String
Dim lsFechaHasta As String
Dim Lineas As Long
Dim TotalDebe As Currency
Dim TotalHaber As Currency
Dim TotalImporte As Currency
Dim TotalSust As Currency
Dim TotalRend As Currency

Dim Total As Long
Dim j As Long
Dim lsObjARendir As String
Dim i As Integer
Dim lsFecha As String
Dim lnImporte As Currency
Dim lsCodCtaCont As String
Dim lsHoja As String
Dim nLin As Integer
Dim lsImpre As String
Dim lnPaginas As Integer
Dim lsMsgErr As String
Dim lsFecRend As String
On Error GoTo ErrReporteArendir
Set rs = CargaOpeCta(psOpeCod, "D")

Dim lsCodCtaCont1, lsCodCtaCont2 As String 'Agregado por ELRO el 20120922

If rs.EOF Then
    RSClose rs
    MsgBox "No se definió Cuenta Contable para Analizar Pendiente", vbInformation, "¡Aviso!"
    Exit Sub
End If
lsCodCtaCont = RSMuestraLista(rs)
'lsCodCtaCont = rs!cCtaContCod

If pbCajaChica Then
    lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoCajaChica & "," & gArendirTipoCajaGeneral & "," & gArendirTipoViaticos
Else
    'lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoCajaGeneral & "," & gArendirTipoViaticos
    lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoCajaGeneral
End If

'Add gitu 08/07/2008
lsArchivo = App.path & "\SPOOLER\LibAuxARendir_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    'oBarra.CloseForm Me
    Exit Sub
End If


lsHoja = "LibAuxARendir"

ExcelAddHoja lsHoja, xlLibro, xlHoja1

'xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(1500, 6)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 7), xlHoja1.Cells(1500, 7)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 8), xlHoja1.Cells(1500, 8)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1500, 1)).NumberFormat = "dd/mm/yyyy"
xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(1500, 6)).NumberFormat = "dd/mm/yyyy"


lsImpre = ""
lnPaginas = 0
Linea lsImpre, PrnSet("MI", 5) & ReporteArendirCuentaLibroEncabezado(pdFecha, pdFecha2, Val(Mid(psOpeCod, 3, 1)), lnPaginas), 0

Lineas = 6
lsFechaDesde = Format(pdFecha, "yyyymmdd")
lsFechaHasta = Format(pdFecha2, "yyyymmdd")

lsCodCtaCont2 = "29" & Mid(psOpeCod, 3, 1) & "80706"


'***Modificado por ELRO el 20120922, según TIC1208270004
'sql = "SELECT  M.CMOVNRO, M.CMOVDESC, MC.NMOVIMPORTE IMPORTE, md.nDocTpo, md.cDocNro, " _
'    & "             P.cPersNombre, ISNULL(ME.NMOVMEIMPORTE,0) AS IMPORTEME, " _
'    & "             ISNULL(Max(ISNULL(Rend.cMovNro, Sust.cMovNro)),'') MovRend, ISNULL(SUM(Sust.nMovImporte),0) ImporteSust, ISNULL(SUM(Sust.nMovMEImporte),0) ImporteMESust, " _
'    & "             ISNULL(Rend.nMovImporte,0) ImporteRend, ISNULL(Rend.nMovMEImporte,0) ImporteMERend " _
'    & "     FROM     MOV M LEFT JOIN MOVDOC MD ON MD.nMovNro=M.nMovNro and NOT md.nDocTpo = " & TpoDocVoucherEgreso _
'    & "         INNER JOIN MOVCta MC ON M.nMovNro=MC.nMovNro " _
'    & "         LEFT JOIN MOVME ME  ON ME.nMovNro=MC.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "         INNER JOIN MOVRef MR ON MR.nMovNro=M.nMovNro " _
'    & "         INNER JOIN MOVArendir MO ON MO.nMovNro=MR.nMovNroRef " _
'    & "         INNER JOIN Persona P ON MO.cPersCOD= P.cPersCOD " _
'    & "         LEFT JOIN (SELECT m.cMovNro, mr.nMovNroRef, mc.nMovImporte* -1 nMovImporte, ISNULL(me.nMovMEImporte,0)*-1 nMovMEImporte " _
'    & "                     FROM Mov m JOIN MovCta mc ON m.nMovNro = mc.nMovNro " _
'    & "                         LEFT JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "                         JOIN MovRef mr on mr.nMovNro = m.nMovNro " _
'    & "                     WHERE m.nMovEstado in( " & gMovEstContabMovContable & ") and m.nMovFlag = " & gMovFlagVigente & " and mc.cCtaContCod IN (" & lsCodCtaCont & ", '29" & Mid(psOpeCod, 3, 1) & "80706' ) " _
'    & "                         and not m.cOpeCod LIKE '40__[356]%' " _
'    & "                     ) Sust ON (Sust.nMovNroRef = m.nMovNro and mo.cTpoArendir = 1) or  (Sust.nMovNroRef = mo.nMovNro and mo.cTpoArendir = 2) "
'
'sql = sql & "   LEFT JOIN (SELECT m.cMovNro, mr.nMovNroRef, mc.nMovImporte* -1 nMovImporte, ISNULL(me.nMovMEImporte,0)*-1 nMovMEImporte " _
'    & "                     FROM Mov m JOIN MovCta mc ON m.nMovNro = mc.nMovNro " _
'    & "                         LEFT JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "                         JOIN MovRef mr on mr.nMovNro = m.nMovNro " _
'    & "                     WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente & " and  " _
'    & "                         ((mc.cCtaContCod in (" & lsCodCtaCont & ") and m.cOpeCod LIKE '40__5%') or (mc.cCtaContCod = '29" & Mid(psOpeCod, 3, 1) & "80706' and m.cOpeCod LIKE '40__6%') ) " _
'    & "                     ) Rend ON (Rend.nMovNroRef = m.nMovNro and mo.cTpoArendir = 1) or  (Rend.nMovNroRef = mo.nMovNro and mo.cTpoArendir = 2) " _
'    & "     WHERE MC.CCTACONTCOD in (" & lsCodCtaCont & ") AND MC.NMOVIMPORTE > 0 AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & lsFechaDesde & "' AND '" & lsFechaHasta & "' " _
'    & "         AND M.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND M.cOpeCod LIKE '40__3%' " _
'    & "         AND MO.cTpoArendir in (1) " _
'    & "     GROUP BY M.CMOVNRO, MR.NMOVNROREF, M.CMOVDESC, MC.NMOVIMPORTE, md.nDocTpo, md.cDocNro, " _
'    & "             P.cPersNombre, ISNULL(ME.NMOVMEIMPORTE,0), " _
'    & "             ISNULL(Rend.nMovImporte,0), ISNULL(Rend.nMovMEImporte,0)" _
'    & "ORDER BY m.cMovNro "
sql = "exec stp_sel_ReporteArendirCuentaLibro  " & lsCodCtaCont & ", '" & lsCodCtaCont2 & "', '" & lsFechaDesde & "', '" & lsFechaHasta & "'"
'***Fin Modificado por ELRO el 20120922

TotalImporte = 0
TotalSust = 0
TotalRend = 0
Set oBarra = New clsProgressBar
ProgressShow oBarra, frmReportes, ePCap_CaptionPercent, True
oBarra.Progress 0, "A Rendir Cuenta : Libro Auxiliar", "Cargando datos...", "", vbBlue

Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
nLin = 7
Total = rs.RecordCount
oBarra.Progress 1, "A Rendir Cuenta : Libro Auxiliar", "Cargando datos...", "", vbBlue
oBarra.Max = Total

Do While Not rs.EOF
    oBarra.Progress rs.Bookmark, "A Rendir Cuenta : Libro Auxiliar", "Generando Reporte...", "", vbBlue
    j = j + 1
    lnImporte = IIf(Mid(psOpeCod, 3, 1) = "1", rs!Importe, rs!ImporteME)
    lsFecha = Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4)
    If rs!MovRend = "" Then
        lsFecRend = ImpreFormat("", 10)
    Else
        lsFecRend = ImpreFormat(GetFechaMov(rs!MovRend, True), 10)
    End If
    Linea lsImpre, ImpreFormat(Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), 12, 0) & _
            ImpreFormat(Format(rs!nDocTpo, "00") & " " & rs!cDocNro, 14) & ImpreFormat(PstaNombre(rs!cPersNombre, False), 32) & ImpreFormat(rs!cMovDesc, 40) & _
            IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, ImpreFormat(rs!ImporteME, 10) & _
               lsFecRend & ImpreFormat(rs!ImporteMESust, 10) & ImpreFormat(rs!ImporteMERend, 10), _
               ImpreFormat(rs!Importe, 10) & ImpreFormat(lsFecRend, 10) & _
               ImpreFormat(rs!ImporteSust, 10) & ImpreFormat(rs!ImporteRend, 10))
    
    xlHoja1.Cells(nLin, 1) = lsFecha
    xlHoja1.Cells(nLin, 2) = Format(rs!nDocTpo, "00") & " " & rs!cDocNro
    xlHoja1.Cells(nLin, 3) = PstaNombre(rs!cPersNombre, False)
    xlHoja1.Cells(nLin, 4) = ImpreFormat(rs!cMovDesc, 180)
    xlHoja1.Cells(nLin, 5) = lnImporte
    xlHoja1.Cells(nLin, 6) = lsFecRend
    xlHoja1.Cells(nLin, 7) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteSust, rs!ImporteMESust)
    xlHoja1.Cells(nLin, 8) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteRend, rs!ImporteMERend)
        
    Lineas = Lineas + 1
    If Lineas > 60 Then
        Linea lsImpre, ReporteArendirCuentaLibroEncabezado(pdFecha, pdFecha2, Val(Mid(psOpeCod, 3, 1)), lnPaginas), 0
        Lineas = 6
    End If
    TotalImporte = TotalImporte + rs!Importe
    TotalSust = TotalSust + rs!ImporteSust
    TotalRend = TotalRend + rs!ImporteRend
    nLin = nLin + 1
    rs.MoveNext
    DoEvents
Loop
    
xlHoja1.Cells(nLin + 1, 4) = "TOTALES :"
xlHoja1.Cells(nLin + 1, 4).HorizontalAlignment = xlHAlignRight
xlHoja1.Cells(nLin + 1, 5) = TotalImporte
xlHoja1.Cells(nLin + 1, 7) = TotalSust
xlHoja1.Cells(nLin + 1, 8) = TotalRend
    
RSClose rs
ProgressClose oBarra, frmReportes, True
Set oBarra = Nothing
Linea lsImpre, String(154, "=")
Linea lsImpre, ImpreFormat("TOTALES :", 10, 89) & ImpreFormat(TotalImporte, 15, 2) & space(11) & ImpreFormat(TotalSust, 11, 2) & ImpreFormat(TotalRend, 10, 2)
If lsImpre <> "" Then
    EnviaPrevio lsImpre, "A RENDIR CUENTA : LIBRO AUXILIAR", gnLinPage, True
End If

ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.path & "\SPOOLER\"

Exit Sub
ErrReporteArendir:
lsMsgErr = Err.Description
    ProgressClose oBarra, frmReportes, True
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ReporteArendirCuentaLibroEncabezado(pdFecha As Date, pdFecha2 As Date, pnMoneda As Integer, Paginas As Integer)
Dim lsCabe As String
lsCabe = CabeRepo(gsNomCmac, gsNomAge, "CAJA GENERAL", IIf(pnMoneda = 1, "SOLES", "DOLARES"), Format(gdFecSis, gsFormatoFechaView), "REPORTE DE A RENDIR A CUENTA", "Desde " & Format(pdFecha, gsFormatoFechaView) & " Hasta " & Format(pdFecha2, gsFormatoFechaView), "", "", _
                      Paginas, 150)

Linea lsCabe, String(150, "-")
Linea lsCabe, ImpreFormat("FECHA", 12, 0) & ImpreFormat("DOCUMENTO", 15) & ImpreFormat("PERSONA", 35) & ImpreFormat("DESCRIPCION", 40) & ImpreFormat("IMPORTE", 10) & ImpreFormat("FEC.REND.", 11) & ImpreFormat("SUSTENTADO", 10) & ImpreFormat("RENDIDO", 10) & Chr(10)
Linea lsCabe, String(150, "-")

'Add By Gitu 08/07/2004

xlHoja1.Range("B1:I1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2
xlHoja1.Range("A6:Z1").EntireColumn.Font.Size = 7
'xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
xlHoja1.Range("B1:B2").RowHeight = 17
xlHoja1.Range("A1:A1").ColumnWidth = 10
xlHoja1.Range("B1:B1").ColumnWidth = 10
xlHoja1.Range("C1:C1").ColumnWidth = 50
xlHoja1.Range("D1:D1").ColumnWidth = 100
xlHoja1.Range("E1:E1").ColumnWidth = 12
xlHoja1.Range("F1:F1").ColumnWidth = 10
xlHoja1.Range("G1:G1").ColumnWidth = 12
xlHoja1.Range("H1:H1").ColumnWidth = 12
'xlHoja1.Range("I1:I1").ColumnWidth = 12
    
xlHoja1.Range("A2:B2").Font.Size = 12
xlHoja1.Range("A3:H5").Font.Size = 10
xlHoja1.Range("A1:H4").Font.Bold = True
xlHoja1.Cells(2, 2) = "L I B R O  A U X I L I A R  D E  A  R E N D I R"
xlHoja1.Cells(3, 3) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
xlHoja1.Cells(4, 1) = "INSTITUCION : " & gsNomCmac
xlHoja1.Range("A2:H2").Merge True
xlHoja1.Range("A3:H3").Merge True
xlHoja1.Range("A2:H3").HorizontalAlignment = xlHAlignCenter
    
xlHoja1.Range("A6:H6").Font.Bold = True
xlHoja1.Cells(6, 1) = "FECHA"
xlHoja1.Cells(6, 2) = "DOCUMENTO"
xlHoja1.Cells(6, 3) = "PERSONA"
xlHoja1.Cells(6, 4) = "DESCRIPCION"
'xlHoja1.Cells(6, 5) = "DIAS"
xlHoja1.Cells(6, 5) = "IMPORTE"
xlHoja1.Cells(6, 6) = "FEC.REND."
xlHoja1.Cells(6, 7) = "SUSTENTADO"
xlHoja1.Cells(6, 8) = "RENDIDO"

xlHoja1.Range("A6:H6").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A6:H6").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A6:H6").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A6:H6").Borders(xlInsideVertical).Color = vbBlack
'xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous


With xlHoja1.PageSetup
     .LeftHeader = ""
     .CenterHeader = ""
     .RightHeader = ""
     .LeftFooter = ""
     .CenterFooter = ""
     .RightFooter = ""

     .PrintHeadings = False
     .PrintGridlines = False
     .PrintComments = xlPrintNoComments
     .CenterHorizontally = True
     .CenterVertically = False
     .Orientation = xlLandscape
     .Draft = False
     .FirstPageNumber = xlAutomatic
     .Order = xlDownThenOver
     .BlackAndWhite = False
     .Zoom = 55
End With

ReporteArendirCuentaLibroEncabezado = lsCabe
End Function
'TORE - Automatizacion de Prorrogas
'Se modifico lel orden de las columnas.
Private Function ReporteArendirCuentaViaticosLibroEncabezado(pdFecha As Date, pdFecha2 As Date, pnMoneda As Integer, Paginas As Integer)
Dim lsCabe As String
lsCabe = CabeRepo(gsNomCmac, gsNomAge, "CAJA GENERAL", IIf(pnMoneda = 1, "SOLES", "DOLARES"), Format(gdFecSis, gsFormatoFechaView), "REPORTE DE A RENDIR A CUENTA VIATICOS", "Desde " & Format(pdFecha, gsFormatoFechaView) & " Hasta " & Format(pdFecha2, gsFormatoFechaView), "", "", _
                      Paginas, 150)
                      
Linea lsCabe, String(220, "-")
Linea lsCabe, ImpreFormat("FECHA", 12, 0) & ImpreFormat("DOCUMENTO", 15) & ImpreFormat("PERSONA", 35) & ImpreFormat("DESCRIPCION", 40) & ImpreFormat("DIAS", 10) & ImpreFormat("PLAZO A RENDIR", 14) & ImpreFormat("PRÓRROGA", 14) & ImpreFormat("DIAS ATRASO", 11) & ImpreFormat("IMPORTE", 10) & ImpreFormat("FEC.REND.", 11) & ImpreFormat("SUSTENTADO", 10) & ImpreFormat("RENDIDO", 10) & Chr(10)
Linea lsCabe, String(220, "-")

xlHoja1.Range("B1:K1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2
xlHoja1.Range("A6:Z1").EntireColumn.Font.Size = 7
'xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
'Modificado PASI20140512 TI-ERS060-2014
'xlHoja1.Range("B1:B2").RowHeight = 17
'xlHoja1.Range("A1:A1").ColumnWidth = 8
'xlHoja1.Range("B1:B1").ColumnWidth = 10
'xlHoja1.Range("C1:C1").ColumnWidth = 50
'xlHoja1.Range("D1:D1").ColumnWidth = 100
'xlHoja1.Range("E1:E1").ColumnWidth = 10
'
'xlHoja1.Range("F1:F1").ColumnWidth = 15
'xlHoja1.Range("G1:G1").ColumnWidth = 10
'
'xlHoja1.Range("H1:H1").ColumnWidth = 12
'xlHoja1.Range("I1:I1").ColumnWidth = 10
'xlHoja1.Range("J1:J1").ColumnWidth = 12
'xlHoja1.Range("K1:K1").ColumnWidth = 12
'
'
'xlHoja1.Range("A2:B2").Font.Size = 12
'xlHoja1.Range("A3:K5").Font.Size = 10
'xlHoja1.Range("A1:K4").Font.Bold = True
'xlHoja1.Cells(2, 2) = "L I B R O  A U X I L I A R  D E  V I A T I C O S"
'xlHoja1.Cells(3, 3) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
'xlHoja1.Cells(4, 1) = "INSTITUCION : " & gsNomCmac
'xlHoja1.Range("A2:K2").Merge True
'xlHoja1.Range("A3:K3").Merge True
'xlHoja1.Range("A2:K3").HorizontalAlignment = xlHAlignCenter
'
'xlHoja1.Range("A6:K6").Font.Bold = True
'xlHoja1.Cells(6, 1) = "FECHA"
'xlHoja1.Cells(6, 2) = "DOCUMENTO"
'xlHoja1.Cells(6, 3) = "PERSONA"
'xlHoja1.Cells(6, 4) = "DESCRIPCION"
'xlHoja1.Cells(6, 5) = "DIAS"
'
'xlHoja1.Cells(6, 6) = "PLAZO A RENDIR"
'xlHoja1.Cells(6, 7) = "DIAS ATRASO"
'
'xlHoja1.Cells(6, 8) = "IMPORTE"
'xlHoja1.Cells(6, 9) = "FEC.REND."
'xlHoja1.Cells(6, 10) = "SUSTENTADO"
'xlHoja1.Cells(6, 11) = "RENDIDO"
'
'xlHoja1.Range("A6:K6").HorizontalAlignment = xlHAlignCenter
'xlHoja1.Range("A6:K6").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
'xlHoja1.Range("A6:K6").Borders(xlInsideVertical).LineStyle = xlContinuous
'xlHoja1.Range("A6:K6").Borders(xlInsideVertical).Color = vbBlack
''xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
''xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous

xlHoja1.Range("B1:B2").RowHeight = 17
xlHoja1.Range("A1:A1").ColumnWidth = 15
xlHoja1.Range("B1:B1").ColumnWidth = 10
xlHoja1.Range("C1:C1").ColumnWidth = 50
xlHoja1.Range("D1:D1").ColumnWidth = 100
xlHoja1.Range("E1:E1").ColumnWidth = 15

xlHoja1.Range("F1:F1").ColumnWidth = 10
xlHoja1.Range("G1:G1").ColumnWidth = 15

xlHoja1.Range("H1:H1").ColumnWidth = 15
xlHoja1.Range("I1:I1").ColumnWidth = 15
xlHoja1.Range("J1:J1").ColumnWidth = 12
xlHoja1.Range("K1:K1").ColumnWidth = 12
xlHoja1.Range("L1:L1").ColumnWidth = 12
xlHoja1.Range("M1:M1").ColumnWidth = 12
xlHoja1.Range("N1:N1").ColumnWidth = 12

    
xlHoja1.Range("A2:B2").Font.Size = 12
xlHoja1.Range("A3:M5").Font.Size = 10
xlHoja1.Range("A1:L4").Font.Bold = True
xlHoja1.Cells(2, 2) = "L I B R O  A U X I L I A R  D E  V I A T I C O S"
xlHoja1.Cells(3, 3) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
xlHoja1.Cells(4, 1) = "INSTITUCION : " & gsNomCmac
xlHoja1.Range("A2:M2").Merge True
xlHoja1.Range("A3:M3").Merge True
xlHoja1.Range("A2:M3").HorizontalAlignment = xlHAlignCenter
    
xlHoja1.Range("A6:M6").Font.Bold = True
xlHoja1.Cells(6, 1) = "FECHA DE SOLICITUD"
xlHoja1.Cells(6, 2) = "DOCUMENTO"
xlHoja1.Cells(6, 3) = "PERSONA"
xlHoja1.Cells(6, 4) = "DESCRIPCION"
xlHoja1.Cells(6, 5) = "FECHA DE PARTIDA"
xlHoja1.Cells(6, 6) = "DIAS"

xlHoja1.Cells(6, 7) = "FECHA DE LLEGADA"
xlHoja1.Cells(6, 8) = "PLAZO A RENDIR"
xlHoja1.Cells(6, 9) = "PRÓRROGA"
xlHoja1.Cells(6, 10) = "DIAS ATRASO"

xlHoja1.Cells(6, 11) = "IMPORTE"
xlHoja1.Cells(6, 12) = "FEC.REND."
xlHoja1.Cells(6, 13) = "SUSTENTADO"
xlHoja1.Cells(6, 14) = "RENDIDO"

xlHoja1.Range("A6:N6").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A6:N6").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A6:N6").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A6:N6").Borders(xlInsideVertical).Color = vbBlack
'xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous
'end PASI

With xlHoja1.PageSetup
     .LeftHeader = ""
     .CenterHeader = ""
     .RightHeader = ""
     .LeftFooter = ""
     .CenterFooter = ""
     .RightFooter = ""

     .PrintHeadings = False
     .PrintGridlines = False
     .PrintComments = xlPrintNoComments
     .CenterHorizontally = True
     .CenterVertically = False
     .Orientation = xlLandscape
     .Draft = False
     .FirstPageNumber = xlAutomatic
     .Order = xlDownThenOver
     .BlackAndWhite = False
     .Zoom = 55
End With

ReporteArendirCuentaViaticosLibroEncabezado = lsCabe
End Function

Public Sub ReporteARendirCuentaPendientes(ByVal psOpeCod As String, ByVal pnMoneda As Moneda, ByVal pdFecha As Date)
'ReporteSustentacionARendirCuenta
On Error GoTo GeneraRepEntregaARendirErr
Dim rs As ADODB.Recordset
Dim lsCtaCod As String
Dim lsCtaContDev As String
Dim lsTipoDoc As String
Dim lsObjARendir As String
Dim lsImpre As String
Dim nTot As Currency
Dim nSdo As Currency
Dim lnLin As Integer
Dim lnPag As Integer
Dim lnRendir As Integer


lnRendir = 1

Set rs = CargaOpeCta(psOpeCod, "D")
If rs.EOF Then
    RSClose rs
    MsgBox "No se definió Cuenta Contable para Analizar Pendiente", vbInformation, "¡Aviso!"
    Exit Sub
End If
lsCtaCod = RSMuestraLista(rs, 0)
lsCtaCod = Mid(lsCtaCod, 2)
lsCtaCod = Left(lsCtaCod, Len(lsCtaCod) - 1)
lsCtaContDev = ""

Set oBarra = New clsProgressBar
ProgressShow oBarra, frmReportes, ePCap_CaptionPercent, True
oBarra.Progress 0, "A Rendir Cuenta : Pendientes", "Cargando datos...", "", vbBlue
Dim oARend As New NARendir
Set rs = oARend.ARendirPendientesTotal(pnMoneda, lsCtaCod, pdFecha, False)
Set oARend = Nothing
If Not rs Is Nothing Then
oBarra.Progress 1, "A Rendir Cuenta : Pendientes", "Cargando datos...", "", vbBlue
If rs.State = adStateOpen Then
   If Not rs.EOF Then
      lsCtaCod = ""
      oBarra.Max = rs.RecordCount
      rs.MoveFirst
      lnLin = gnLinPage
      Do While Not rs.EOF
         If lnLin > gnLinPage - 4 Then
            Linea lsImpre, CabeRepo(gsNomCmac, gsNomAge, "CAJA GENERAL", IIf(pnMoneda = 1, "SOLES", "DOLARES"), Format(gdFecSis, gsFormatoFechaView), "REPORTE DE A RENDIR A CUENTA : PENDIENTES ", " AL " & Format(pdFecha, gsFormatoFechaView), "", "", lnPag, 130), 0
            Linea lsImpre, " FECHA              CONCEPTO                                  IMPORTE          SALDO    PERSONA                        DOCUMENTO "
            Linea lsImpre, String(130, "-")
            lnLin = 7
         End If
         oBarra.Progress rs.Bookmark, "A Rendir Cuenta : Pendientes", "Generando Reporte...", "", vbBlue
'JEOM Comentado en Iquitos se usa la misma cuenta para viaticos y A Rendir
'         If lsCtaCod <> rs!cCtaContCod Then
'            If lsCtaCod <> "" Then
'                Linea lsImpre, String(130, "-")
'                Linea lsImpre, "TOTALES : " & PrnVal(nTot, 16, 2) & " " & PrnVal(nSdo, 16, 2)
'                nTot = 0
'                nSdo = 0
'            End If
'            Linea lsImpre, ""
'            Linea lsImpre, "Cuenta : " & rs!cCtaContCod & " - "  '& rs!cCtaContDesc
'            lsCtaCod = rs!cCtaContCod
'         End If

'JEOM 17-05-07
         If lnRendir <> rs!Rendir Then
            Linea lsImpre, String(130, "-")
            Linea lsImpre, space(46) & "TOTALES : " & PrnVal(nTot, 14, 2) & " " & PrnVal(nSdo, 14, 2)
            nTot = 0
            nSdo = 0
            
            Linea lsImpre, ""
            Linea lsImpre, CabeRepo(gsNomCmac, gsNomAge, "CAJA GENERAL", IIf(pnMoneda = 1, "SOLES", "DOLARES"), Format(gdFecSis, gsFormatoFechaView), "REPORTE DE A RENDIR A CUENTA VIATICOS: PENDIENTES ", " AL " & Format(pdFecha, gsFormatoFechaView), "", "", lnPag, 130), 0
            lnRendir = rs!Rendir
            lnLin = 5
         End If
'FIN JEOM
         
         Linea lsImpre, Left(rs!cMovNro, 8) & "-" & Mid(rs!cMovNro, 9, 6) & Right(rs!cMovNro, 4) & " ", 0
         Linea lsImpre, Justifica(rs!cMovDesc, 35) & " ", 0
         Linea lsImpre, PrnVal(rs!nMovImporte, 14, 2) & " ", 0
         Linea lsImpre, PrnVal(rs!nSaldo, 14, 2) & "  ", 0
         Linea lsImpre, Justifica(IIf(IsNull(rs!cPErsona), "", rs!cPErsona), 30) & " ", 0
         Linea lsImpre, rs!cDocAbrev & " " & rs!cDocNro
         nTot = nTot + rs!nMovImporte
         nSdo = nSdo + rs!nSaldo
         lnLin = lnLin + 1
         rs.MoveNext
      Loop
      Linea lsImpre, String(130, "-")
      Linea lsImpre, space(46) & "TOTALES : " & PrnVal(nTot, 14, 2) & " " & PrnVal(nSdo, 14, 2)
   Else
      MsgBox "No existen Pendientes de Regularización", vbInformation, "¡Aviso!"
   End If
End If
End If
RSClose rs
ProgressClose oBarra, frmReportes, True
Set oBarra = Nothing

If lsImpre <> "" Then
    EnviaPrevio lsImpre, "A Rendir Cuenta: Pendiente", gnLinPage, True
End If

Exit Sub
GeneraRepEntregaARendirErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Public Function CabeRepo2(ByVal nCarLin As Integer, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date) As String

Dim sTit1 As String, sTit2 As String
Dim sCabe01 As String, sCabe02 As String
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & space((nCarLin - 36) - (Len(sCabe01) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & space(5) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & space((nCarLin - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & "HORA :   " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
sTit1 = String(Int((nCarLin - Len(sTitRp1)) / 2), " ") & sTitRp1
sTit2 = String(Int((nCarLin - Len(sTitRp2)) / 2), " ") & sTitRp2
    
CabeRepo2 = CabeRepo2 & sCabe01 & Chr$(10)
CabeRepo2 = CabeRepo2 & sCabe02 & Chr$(10)
CabeRepo2 = CabeRepo2 & sTit1 & Chr$(10)
CabeRepo2 = CabeRepo2 & sTit2
End Function

Private Sub CadeceraRepLinFinDescalce(ByVal nTCF As Double, ByVal dFecha As Date)
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.LeftMargin = 1.5
xlHoja1.PageSetup.RightMargin = 1.5
xlHoja1.PageSetup.TopMargin = 1
xlHoja1.PageSetup.BottomMargin = 1
xlHoja1.Cells(1, 1) = "CAJA MUNICIPAL DE MAYNAS S.A"
xlHoja1.Cells(2, 1) = UCase(gsNomAge)
xlHoja1.Cells(2, 7) = "FECHA"
xlHoja1.Cells(2, 8) = gdFecSis
xlHoja1.Cells(2, 9) = Time
xlHoja1.Cells(3, 1) = "REPORTE SALDOS LINEAS DE FINANCIAMIENTO Y DESCALCE"
xlHoja1.Cells(4, 1) = "AL " & Format(dFecha, "dd/mm/yyyy") & " ( En Nuevos Soles )"
xlHoja1.Cells(5, 1) = "Tipo Cambio Fijo : "
xlHoja1.Cells(5, 2) = Format$(nTCF, "#,##0.0000")

xlHoja1.Range("A3:M3").Merge
xlHoja1.Range("A4:M4").Merge
xlHoja1.Range("A3:M4").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A1:M3").Font.Size = 8

xlHoja1.Cells(6, 2) = "TOTAL COLOCACIONES"
xlHoja1.Cells(6, 8) = "SALDO LINEAS FINANCIAMIENTO"
xlHoja1.Cells(7, 1) = "LINEA FINANCIAMIENTO"
xlHoja1.Cells(7, 2) = "MN NUM"
xlHoja1.Cells(7, 3) = "MN IMPORTE"
xlHoja1.Cells(7, 4) = "ME NUM"
xlHoja1.Cells(7, 5) = "ME IMPORTE"
xlHoja1.Cells(7, 6) = "TOTAL CTAS"
xlHoja1.Cells(7, 7) = "TOTAL MN"
xlHoja1.Cells(7, 8) = "SALD LINEA MN"
xlHoja1.Cells(7, 9) = "SALD LINEA ME"
xlHoja1.Cells(7, 10) = "SALD TOTAL MN"
xlHoja1.Cells(7, 11) = "DESCALCE MN"
xlHoja1.Cells(7, 12) = "DESCALCE ME"
xlHoja1.Cells(7, 13) = "SALD DESC MN"

xlHoja1.Range("B6:G6").Merge
xlHoja1.Range("H6:J6").Merge
xlHoja1.Range("A6:A7").Merge
xlHoja1.Range("K6:K7").Merge
xlHoja1.Range("L6:L7").Merge
xlHoja1.Range("M6:M7").Merge
xlHoja1.Range("A6:M7").HorizontalAlignment = xlHAlignCenter

xlHoja1.Range("A1:M7").Font.Bold = True

xlHoja1.Range("A1:A1").ColumnWidth = 40
xlHoja1.Range("B1:M1").ColumnWidth = 15
xlHoja1.Range("B1:B1").ColumnWidth = 8
xlHoja1.Range("D1:D1").ColumnWidth = 8
xlHoja1.Range("F1:F1").ColumnWidth = 8

xlHoja1.Range("A1:M200").Font.Name = "Verdana"
xlHoja1.Range("A4:M200").Font.Size = 8
xlHoja1.Range("A6:M7").BorderAround xlContinuous
xlHoja1.Range("A6:M7").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("A6:M7").Borders(xlInsideVertical).LineStyle = xlContinuous
End Sub

Public Sub ReporteSaldosLineaFinanciamientoDescalce(ByVal dFecha As Date)
Dim rsAdeud As Recordset
Dim sSql As String
Dim oData As DConecta
Dim nTipoCambio As Double
Dim oTC As nTipoCambio
Dim sRangoTotal As String
Dim sLineaDesc As String, sLinea As String
Dim nSaldoAdeudMN As Double, nSaldoAdeudME As Double
Dim nFilIni As Integer, nFilFin As Integer

dFecha = DateAdd("d", -1, dFecha)

ConsolidaDescalce dFecha

sSql = "Select L.cLineaDesc, K.cConsDescripcion cProductoDesc, L.cLineaCred, L.nProducto, nNumLineaMN, nNumLineaME, " _
    & "L.nSaldoLineaMN, L.nSaldoLineaME, ISNULL(A.nSaldoAdeudMN,0) nSaldoAdeudMN, ISNULL(A.nSaldoAdeudME,0) nSaldoAdeudME From " _
    & "(Select L.cDescripcion cLineaDesc, LEFT(L.cLineaCred,5) cLineaCred, LC.nProducto, " _
    & "nSaldoLineaMN = ISNULL(SUM(CASE WHEN RIGHT(LC.cLineaCred,1) = '1' THEN LC.nSaldo END),0), " _
    & "nSaldoLineaME = ISNULL(SUM(CASE WHEN RIGHT(LC.cLineaCred,1) = '2' THEN LC.nSaldo END),0), " _
    & "nNumLineaMN = ISNULL(SUM(CASE WHEN RIGHT(LC.cLineaCred,1) = '1' THEN LC.nNumero END),0), " _
    & "nNumLineaME = ISNULL(SUM(CASE WHEN RIGHT(LC.cLineaCred,1) = '2' THEN LC.nNumero END),0) " _
    & "From ColocLineaCredito L INNER JOIN ColLineaSaldoAdeud LC ON L.cLineaCred = LC.cLineaCred " _
    & "Where LC.dFecha >= '" & Format$(dFecha, "mm/dd/yyyy") & "' And LC.dFecha < '" & Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy") & "' " _
    & "Group by L.cDescripcion, LEFT(L.cLineaCred,5), LC.nProducto) L LEFT JOIN " _
    & "(Select LEFT(cLineaCred,5) cLineaCred, " _
    & "nSaldoAdeudMN = ISNULL(SUM(CASE WHEN RIGHT(cLineaCred,1) = '1' THEN nSaldo END),0), " _
    & "nSaldoAdeudME = ISNULL(SUM(CASE WHEN RIGHT(cLineaCred,1) = '2' THEN nSaldo END),0) " _
    & "From ColAdeudadoSaldo Where dFecha >= '" & Format$(dFecha, "mm/dd/yyyy") & "' And dFecha < '" & Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy") & "' " _
    & "Group by LEFT(cLineaCred,5)) A  ON L.cLineaCred = A.cLineaCred " _
    & "JOIN Constante K ON L.nProducto = K.nCOnsValor Where K.nConsCod = 3034 " _
    & "Order by L.cLineaCred, L.nProducto"

Set oData = New DConecta
oData.AbreConexion
Set rsAdeud = oData.CargaRecordSet(sSql)
oData.CierraConexion
Set oData = Nothing

If Not (rsAdeud.EOF And rsAdeud.BOF) Then
    Set oTC = New nTipoCambio
    nTipoCambio = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    Set oTC = Nothing
    
    Dim nCol As Integer
    Dim sCol As String, lsArchivo  As String, lsRuta As String
    Dim lbLibroOpen As Boolean

    On Error GoTo ErrImprime

    lsRuta = App.path & "\Spooler\"
    lsArchivo = lsRuta & "LINFIN_DESCALCE" & "_" & Format$(dFecha, "yyyymmdd")
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbLibroOpen Then
        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja Format$(dFecha, "dd mmm yyyy"), xlLibro, xlHoja1
        CadeceraRepLinFinDescalce nTipoCambio, dFecha
        nFilIni = 8
        nFilFin = 8
        sLineaDesc = ""
        sRangoTotal = ""
        sLinea = ""
        Do While Not rsAdeud.EOF
            If sLinea <> rsAdeud("cLineaDesc") Then
                If Trim(sLineaDesc) <> "" Then
                    nFilFin = nFilFin + 1
                    xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).Formula = "=SUM(" & "B" & nFilIni & ":" & "B" & nFilFin - 1 & ")"
                    xlHoja1.Range("C" & nFilFin & ":" & "C" & nFilFin).Formula = "=SUM(" & "C" & nFilIni & ":" & "C" & nFilFin - 1 & ")"
                    xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).Formula = "=SUM(" & "D" & nFilIni & ":" & "D" & nFilFin - 1 & ")"
                    xlHoja1.Range("E" & nFilFin & ":" & "E" & nFilFin).Formula = "=SUM(" & "E" & nFilIni & ":" & "E" & nFilFin - 1 & ")"
                    xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).Formula = "=SUM(" & "F" & nFilIni & ":" & "F" & nFilFin - 1 & ")"
                    xlHoja1.Range("G" & nFilFin & ":" & "G" & nFilFin).Formula = "=SUM(" & "G" & nFilIni & ":" & "G" & nFilFin - 1 & ")"
                    xlHoja1.Cells(nFilFin, 8) = nSaldoAdeudMN
                    xlHoja1.Cells(nFilFin, 9) = nSaldoAdeudME
                    xlHoja1.Range("J" & nFilFin & ":" & "J" & nFilFin).Formula = "=+H" & nFilFin & " + I" & nFilFin & "*$B$5"
                    xlHoja1.Range("K" & nFilFin & ":" & "K" & nFilFin).Formula = "=+H" & nFilFin & " - C" & nFilFin
                    xlHoja1.Range("L" & nFilFin & ":" & "L" & nFilFin).Formula = "=+I" & nFilFin & " - E" & nFilFin
                    xlHoja1.Range("M" & nFilFin & ":" & "M" & nFilFin).Formula = "=+K" & nFilFin & " + L" & nFilFin & "*$B$5"
                    xlHoja1.Cells(nFilFin, 1) = "TOTAL " & Trim(sLineaDesc)
                    xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Font.Bold = True
                    xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).BorderAround xlContinuous
                    xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Borders(xlInsideVertical).LineStyle = xlContinuous
                    xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).NumberFormat = "#,##0.00"
                    xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).NumberFormat = "#,##0"
                    xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).NumberFormat = "#,##0"
                    xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).NumberFormat = "#,##0"
                    
                    nFilIni = nFilFin + 2
                    nFilFin = nFilIni
                End If
                sLineaDesc = Trim(rsAdeud("cLineaDesc"))
                sLinea = Trim(rsAdeud("cLineaDesc"))
                xlHoja1.Cells(nFilFin, 1) = sLineaDesc
                xlHoja1.Range("A" & nFilFin & ":" & "I" & nFilFin).Font.Bold = True
                nFilFin = nFilFin + 1
                xlHoja1.Cells(nFilFin, 1) = space(2) & rsAdeud("cProductoDesc")
                xlHoja1.Cells(nFilFin, 2) = rsAdeud("nNumLineaMN")
                xlHoja1.Cells(nFilFin, 3) = rsAdeud("nSaldoLineaMN")
                xlHoja1.Cells(nFilFin, 4) = rsAdeud("nNumLineaME")
                xlHoja1.Cells(nFilFin, 5) = rsAdeud("nSaldoLineaME")
                nSaldoAdeudMN = rsAdeud("nSaldoAdeudMN")
                nSaldoAdeudME = rsAdeud("nSaldoAdeudME")
                xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).Formula = "=+B" & nFilFin & " + D" & nFilFin
                xlHoja1.Range("G" & nFilFin & ":" & "G" & nFilFin).Formula = "=+C" & nFilFin & " + E" & nFilFin & "*$B$5"
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).BorderAround xlContinuous
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Borders(xlInsideVertical).LineStyle = xlContinuous
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).NumberFormat = "#,##0.00"
                xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).NumberFormat = "#,##0"
                xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).NumberFormat = "#,##0"
                xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).NumberFormat = "#,##0"
                
            Else
                nFilFin = nFilFin + 1
                xlHoja1.Cells(nFilFin, 1) = space(2) & rsAdeud("cProductoDesc")
                xlHoja1.Cells(nFilFin, 1) = space(2) & rsAdeud("cProductoDesc")
                xlHoja1.Cells(nFilFin, 2) = rsAdeud("nNumLineaMN")
                xlHoja1.Cells(nFilFin, 3) = rsAdeud("nSaldoLineaMN")
                xlHoja1.Cells(nFilFin, 4) = rsAdeud("nNumLineaME")
                xlHoja1.Cells(nFilFin, 5) = rsAdeud("nSaldoLineaME")
                nSaldoAdeudMN = rsAdeud("nSaldoAdeudMN")
                nSaldoAdeudME = rsAdeud("nSaldoAdeudME")
                xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).Formula = "=+B" & nFilFin & " + D" & nFilFin
                xlHoja1.Range("G" & nFilFin & ":" & "G" & nFilFin).Formula = "=+C" & nFilFin & " + E" & nFilFin & "*$B$5"
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).BorderAround xlContinuous
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Borders(xlInsideVertical).LineStyle = xlContinuous
                xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).NumberFormat = "#,##0.00"
                xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).NumberFormat = "#,##0"
                xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).NumberFormat = "#,##0"
                xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).NumberFormat = "#,##0"
            End If
            rsAdeud.MoveNext
        Loop
        nFilFin = nFilFin + 1
        rsAdeud.MoveLast
        xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).Formula = "=SUM(" & "B" & nFilIni & ":" & "B" & nFilFin - 1 & ")"
        xlHoja1.Range("C" & nFilFin & ":" & "C" & nFilFin).Formula = "=SUM(" & "C" & nFilIni & ":" & "C" & nFilFin - 1 & ")"
        xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).Formula = "=SUM(" & "D" & nFilIni & ":" & "D" & nFilFin - 1 & ")"
        xlHoja1.Range("E" & nFilFin & ":" & "E" & nFilFin).Formula = "=SUM(" & "E" & nFilIni & ":" & "E" & nFilFin - 1 & ")"
        xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).Formula = "=SUM(" & "F" & nFilIni & ":" & "F" & nFilFin - 1 & ")"
        xlHoja1.Range("G" & nFilFin & ":" & "G" & nFilFin).Formula = "=SUM(" & "G" & nFilIni & ":" & "G" & nFilFin - 1 & ")"
        xlHoja1.Cells(nFilFin, 8) = nSaldoAdeudMN
        xlHoja1.Cells(nFilFin, 9) = nSaldoAdeudME
        xlHoja1.Range("J" & nFilFin & ":" & "J" & nFilFin).Formula = "=+H" & nFilFin & " + I" & nFilFin & "*$B$5"
        xlHoja1.Range("K" & nFilFin & ":" & "K" & nFilFin).Formula = "=+H" & nFilFin & " - C" & nFilFin
        xlHoja1.Range("L" & nFilFin & ":" & "L" & nFilFin).Formula = "=+I" & nFilFin & " - E" & nFilFin
        xlHoja1.Range("M" & nFilFin & ":" & "M" & nFilFin).Formula = "=+K" & nFilFin & " + L" & nFilFin & "*$B$5"
        xlHoja1.Cells(nFilFin, 1) = "TOTAL " & Trim(sLineaDesc)
        xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Font.Bold = True
        xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).BorderAround xlContinuous
        xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("A" & nFilFin & ":" & "M" & nFilFin).NumberFormat = "#,##0.00"
        xlHoja1.Range("B" & nFilFin & ":" & "B" & nFilFin).NumberFormat = "#,##0"
        xlHoja1.Range("D" & nFilFin & ":" & "D" & nFilFin).NumberFormat = "#,##0"
        xlHoja1.Range("F" & nFilFin & ":" & "F" & nFilFin).NumberFormat = "#,##0"
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        CargaArchivo lsArchivo & ".xls", lsRuta
    End If
    
Else
    MsgBox "No Existen datos para este reporte", vbInformation, "Aviso"
End If
rsAdeud.Close
Set rsAdeud = Nothing
Exit Sub
ErrImprime:
   MsgBox Err.Description, vbInformation, "!Aviso!"
   
End Sub

Public Sub ConsolidaDescalce(pdFecha As Date)

Dim oCon As New DConecta
Dim oAge As New DActualizaDatosArea
Dim sRuta As String
Dim nBan As Boolean
Dim sNombreTabla As String
Dim sNombreTabla2 As String

On Error GoTo Err_
Dim sSql As String
'Dim rAge As New ADODB.Recordset

Dim rTemp As New ADODB.Recordset
Dim sSqlT As String

 
Dim nIndiceVac As Double

    
    sNombreTabla = "ColLineaSaldoAdeud"
    sNombreTabla2 = "ColAdeudadoSaldo"
    
    sSqlT = "DELETE FROM " & sNombreTabla & " WHERE Convert(varchar(8), dfecha, 112)='" & Format(pdFecha, "YYYYMMdd") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
        
    sSql = "Insert Into " & sRuta & sNombreTabla & " (cAgeCod,dFecha,cLineaCred,nProducto,nSaldo,nNumero) "
    sSql = sSql & " Select cAgeCod, dFecha, Fondo, Producto, SUM(nSaldo) nSaldo, Sum(nNro) nNumero "
    sSql = sSql & " From ( "
    
    sSql = sSql & " Select CEC.cCodAge as cAgeCod, CEC.dEstad as dFecha, "
    sSql = sSql & "    L0.cLineaCred as Fondo,"
    sSql = sSql & "    Substring(L1.cLineaCred,7,3) as Producto,"
    sSql = sSql & "    CEC.nSaldoCap as nSaldo, CEC.nNumSaldos nNro"
    sSql = sSql & " From"
    sSql = sSql & "    ColocLineaCredito L1"
    sSql = sSql & "        Inner Join ColocLineaCredito L0"
    sSql = sSql & "            On L0.cLineaCred = Left(L1.cLineaCred,5)"
    sSql = sSql & "        Inner Join ColocEstadDiaCred CEC on CEC.cLineaCred =L1.cLineaCred"
    sSql = sSql & " Where Convert(varchar(8), CEC.dEstad, 112)='" & Format(pdFecha, "YYYYMMdd") & "'"

    sSql = sSql & " Union "
    sSql = sSql & " Select '00' as cAgeCod, '" & Format(pdFecha, "MM/dd/YYYY") & "', L0.cLineaCred as Fondo, "
    sSql = sSql & " '' as Producto, 0 as nSaldo, 0 as nNro "
    sSql = sSql & " From ColocLineaCredito L0 "
    sSql = sSql & " Where L0.cLineaCred Like '_____' "
    
    sSql = sSql & "         ) A Group by cAgeCod, dFecha, Fondo, Producto  "
            
    oCon.Ejecutar sSql
    
    oCon.CierraConexion
            
    'Indice VAC
    
    sSqlT = "Select nIndiceVac From IndiceVac Where "
    sSqlT = sSqlT & " dIndiceVac IN (Select MAX(dIndiceVac) FRom IndiceVac Where dIndiceVac < DateAdd(dd,1,'" & Format(pdFecha, "YYYY/MM/dd") & "'))"
    oCon.AbreConexion
    Set rTemp = oCon.CargaRecordSet(sSqlT)
    If rTemp.BOF Then
        nIndiceVac = 0
    Else
        nIndiceVac = rTemp!nIndiceVac
    End If
    rTemp.Close
    Set rTemp = Nothing
             
    sSqlT = "DELETE FROM " & sNombreTabla2 & " WHERE Convert(varchar(8), dfecha, 112)='" & Format(pdFecha, "YYYYMMdd") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
    oCon.CierraConexion
     
    sSqlT = " Insert Into " & sNombreTabla2 & " (dFecha,cLineaCred,nSaldo) "
    sSqlT = sSqlT & " Select '" & Format(pdFecha, "YYYY/MM/dd") & "', cCodLinCred, SUM(nSaldoCap) nSaldoCap From ( "
    sSqlT = sSqlT & " SELECT CI.cIFTpo, CI.cPersCod, CI.cCtaIFCod, CI.cCtaIFDesc, CI.dCtaIFAper, dCtaIFVenc, "
    sSqlT = sSqlT & " cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cic.nNroCuota, cic.nInteresPagado, "
    sSqlT = sSqlT & " cic.dVencimiento, Round(CIA.nSaldoCap * CASE WHEN SubString(CI.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' "
    sSqlT = sSqlT & " THEN " & nIndiceVac & " ELSE 1 END,2) nSaldoCap , ISNULL(cia.cCodLinCred,'') cCodLinCred, "
    sSqlT = sSqlT & " ISNULL(L.cDescripcion,'') cDesLinCred "
    sSqlT = sSqlT & " FROM CtaIF CI LEFT JOIN CtaIfAdeudados CIA ON CIA.cIFTpo = CI.cIFTpo And CIA.cPersCod = CI.cPersCod "
    sSqlT = sSqlT & " And CIA.cCtaIFCod = CI.cCtaIFCod JOIN ColocLineaCredito L ON L.cLineaCred = CIA.cCodLinCred "
    sSqlT = sSqlT & " LEFT JOIN CtaIFCalendario CIC ON CIC.cIFTpo = ci.cIFTpo and CIC.cPersCod = ci.cPersCod And "
    sSqlT = sSqlT & " CIC.cCtaIFCod = CI.cCtaIFCod And CIC.cTpoCuota = '2' And CIC.nNroCuota = (SELECT Min(nNroCuota) "
    sSqlT = sSqlT & " FROM CtaIFCalendario cic1 Where cic1.cIFTpo = CIC.cIFTpo And cic1.cPersCod = CIC.cPersCod And "
    sSqlT = sSqlT & " cic1.cCtaIFCod = cic.cCtaIFCod And cic1.cTpoCuota = CIC.cTpoCuota And cEstado = 0) "
    sSqlT = sSqlT & " WHERE ci.cCtaIFEstado IN (1,0) and  ci.cIFTpo+ci.cCtaIFCod LIKE '__05%' "
    sSqlT = sSqlT & " ) A Group by cCodLinCred "
            
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
    oCon.CierraConexion
    
    'cmdReportes.Enabled = True
    
    'MsgBox "Consolidación Finalizada satisfactoriamente", vbInformation, "Aviso"
    
Exit Sub
Err_:

End Sub


Public Sub ImprimeConcentracionFondos(ByVal dFechaAl As Date, ByVal nTipoCambio As Double)

Dim matBancosMN() As String ' [Correlativo del Banco]
Dim matBancosME() As String
 
Dim matCajasMN() As String ' [Correlativo de la Caja]
Dim matCajasME() As String

Dim matArregloBancosMN() As Currency  ' [Correlativo del Banco]
                                   ' [1 to 6] 1 = Ahorro / 2 = Cta Cte / 3 = Plazo Fijo / 4 = Total /  5 = TEA / 6 = Endeudamiento
                                   
Dim matArregloBancosME() As Currency
                                   
Dim matArregloCajasMN() As Currency ' [Correlativo de la Caja]
                                   ' [1 to 6] 1 = Ahorro / 2 = Cta Cte / 3 = Plazo Fijo / 4 = Total /  5 = TEA / 6 = Endeudamiento
Dim nfil As Integer
Dim nFilTemp1 As Integer
Dim nFilTemp2 As Integer
Dim matArregloCajasME() As Currency
                                   
Dim nContCajasMN As Integer
Dim nContCajasME As Integer
Dim nContBancosMN As Integer
Dim nContBancosME As Integer
Dim nPatrEfectivo As Currency
Dim nLiquid(1 To 2) As Currency

Dim sArchConsolBcos(1 To 2) As String '[1=Soles / 2=Dolares] 'Buscar Format(pdFecha, "dd-mm-yyyy")
  
Dim sArchGrabar As String
Dim lbLibroOpen As Boolean

Dim nTemp1 As Integer
Dim nTemp2 As Currency
Dim nValTemp1(1 To 2) As Integer
Dim nValTemp2(1 To 2) As String
Dim nValTemp3(1 To 2) As String
Dim rsTempo As New ADODB.Recordset

Dim lsArchivo As String

Dim bexiste As Boolean
Dim bencontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim nTemp(6) As Currency
Dim nContTemp As Integer
Dim nTem As Integer
Dim nTempTC(1 To 2) As Integer

Dim matTempo(50) As String
Dim ntempo As Integer
Dim nTempo_2 As Integer
Dim K As Integer
Dim nFilaIniRestrin As Integer
Dim nFilaFinRestrin As Integer
    Dim m As Integer
    Dim mTempo As String
    
sArchConsolBcos(1) = App.path & "\Spooler\Anx15A_ConsolBancos_" & Format(dFechaAl, "MMYYYY") & "MN.xls"
sArchConsolBcos(2) = App.path & "\Spooler\Anx15A_ConsolBancos_" & Format(dFechaAl, "MMYYYY") & "ME.xls"
 
sArchGrabar = App.path & "\Spooler\ConcFondos" & Format(dFechaAl, "ddMMYYYY") & ".xls"
 
On Error GoTo ErrBegin

bexiste = fs.FileExists(sArchConsolBcos(1))

If bexiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & sArchConsolBcos(1), vbExclamation, "Aviso!!!"
    Exit Sub
Else
    bexiste = fs.FileExists(sArchConsolBcos(2))
    
    If bexiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & sArchConsolBcos(2), vbExclamation, "Aviso!!!"
        Exit Sub
     
    End If
End If

    'Calculo la liqu en soles y dolares

    Dim oAnx As New NEstadisticas
    nLiquid(1) = Val(oAnx.GetImporteEstadAnexos(dFechaAl, "LIQUIDSOLES", "1"))
    nLiquid(2) = Val(oAnx.GetImporteEstadAnexos(dFechaAl, "LIQUIDDOLARES", "2"))
      
    'Calculo el nPatrEfectivo
    
    nPatrEfectivo = Val(oAnx.GetImporteEstadAnexosMax("01/" & Mid(DateAdd("m", -1, dFechaAl), 4, 7), "TOTALREP03", "1"))
     
    Set oAnx = Nothing
    
    'If nPatrEfectivo = 0 Then
    '    MsgBox "No se ha generado el Reporte 03 para el Calculo del Patrimonio Efectivo", vbInformation, "Aviso!!!"
    'End If
    
    'Abro el Anx15A_ConsolBancos_MMYYYYMN.xls
    'BANCOS MONEDA NACIONAL
    lsArchivo = sArchConsolBcos(1)
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(lsArchivo)

    bencontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If UCase(xlHoja1.Name) = UCase(Format(dFechaAl, "dd-mm-yyyy")) Then
            bencontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bencontrado = False Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        MsgBox "No existen datos con la fecha especificada" & Chr(13) & " en el archivo " & lsArchivo, vbExclamation, "Aviso!!!"
        Exit Sub
    End If

    nTemp(1) = 0
    nTemp(2) = 0
    nTemp(3) = 0
    nTemp(4) = 0
    nTemp(5) = 0
    nTemp(6) = 0
      


'   *************************************
    
    For i = 1 To 50
        matTempo(i) = ""
    Next
    ntempo = 0
    nTempo_2 = 0
    K = 0
    nFilaIniRestrin = 0
    nFilaFinRestrin = 0
    
    i = 9
    nContBancosMN = 0
    Do While UCase(xlHoja1.Cells(i, 1)) <> "CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            
            If Trim(xlHoja1.Cells(i, 1)) = "BANCOS RESTRINGIDOS" Then
                nFilaIniRestrin = i + 1
            ElseIf Trim(xlHoja1.Cells(i, 1)) = "SUBTOTAL BCOS. RESTRINGIDOS" Then
                nFilaFinRestrin = i - 1
            End If
            
            nTempo_2 = 0
            For K = 1 To ntempo
                If Trim(matTempo(K)) = Trim(xlHoja1.Cells(i, 1)) Then
                    nTempo_2 = nTempo_2 + 1
                End If
            Next
            If nTempo_2 = 0 Then
                If Trim(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BANCOS" And Trim(xlHoja1.Cells(i, 1)) <> "BANCOS RESTRINGIDOS" And Trim(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BCOS. RESTRINGIDOS" Then
                    nContBancosMN = nContBancosMN + 1
                    ntempo = ntempo + 1
                    matTempo(ntempo) = Trim(xlHoja1.Cells(i, 1))
                End If
            End If
        End If
        i = i + 1
    Loop
'   *************************************
'    'Saco cantidad de Bancos
'    I = 9
'    Do While UCase(xlHoja1.Cells(I, 1)) <> "SUBTOTAL BANCOS"
'        If Len(Trim(xlHoja1.Cells(I, 1))) > 0 Then
'            nContBancosMN = nContBancosMN + 1
'        End If
'        I = I + 1
'    Loop
'
'    'Saco cantidad de Cajas
'    Do While UCase(xlHoja1.Cells(I, 1)) <> "CMACS"
'        I = I + 1
'    Loop
    
    i = i + 1
    
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            nContCajasMN = nContCajasMN + 1
        End If
        i = i + 1
    Loop
    
    'Bancos
    ReDim matBancosMN(nContBancosMN) As String
    ReDim matArregloBancosMN(nContBancosMN, 6) As Currency
            
    'Cajas
    ReDim matCajasMN(nContCajasMN) As String
    ReDim matArregloCajasMN(nContCajasMN, 5) As Currency
            
    'Bancos
    i = 9
    nContBancosMN = 0
    
    'Obtengo los valores
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BANCOS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            If nContBancosMN > 0 Then
                matArregloBancosMN(nContBancosMN, 1) = nTemp(1)
                matArregloBancosMN(nContBancosMN, 2) = nTemp(2)
                matArregloBancosMN(nContBancosMN, 3) = nTemp(3)
                matArregloBancosMN(nContBancosMN, 4) = nTemp(4)
                matArregloBancosMN(nContBancosMN, 6) = nTemp(6)
                nTemp(1) = 0
                nTemp(2) = 0
                nTemp(3) = 0
                nTemp(4) = 0
                nTemp(6) = 0
            End If
            nContBancosMN = nContBancosMN + 1
            matBancosMN(nContBancosMN) = Trim(xlHoja1.Cells(i, 1))
        End If
        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(i, 5)) ' Ahorro
        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(i, 3)) ' Cta Cte
        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(i, 7)) ' Plazo Fijo
        nTemp(4) = nTemp(4) + Val(xlHoja1.Cells(i, 8)) ' Total
        nTemp(6) = nTemp(6) + Val(xlHoja1.Cells(i, 9)) ' Endeudados
        
        If nFilaIniRestrin = nFilaFinRestrin And nFilaIniRestrin = 0 Then
        Else
            If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
                mTempo = ""
                For m = nFilaIniRestrin To nFilaFinRestrin
                    If Len(Trim(xlHoja1.Cells(m, 1))) > 0 Then
                        mTempo = Trim(xlHoja1.Cells(m, 1))
                    End If
                    If Trim(mTempo) = Trim(matBancosMN(nContBancosMN)) Then
                        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(m, 5)) ' Ahorro
                        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(m, 3)) ' Cta Cte
                        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(m, 7)) ' Plazo Fijo
                        nTemp(4) = nTemp(4) + Val(xlHoja1.Cells(m, 8)) ' Total
                        nTemp(6) = nTemp(6) + Val(xlHoja1.Cells(m, 9)) ' Endeudados
                    End If
                Next
            End If
        End If
        
        i = i + 1
    Loop
    
    If nTemp(1) > 0 Or nTemp(2) > 0 Or nTemp(3) > 0 Or nTemp(4) > 0 Or nTemp(6) > 0 Then
        If nContBancosMN > 0 Then
            matArregloBancosMN(nContBancosMN, 1) = nTemp(1)
            matArregloBancosMN(nContBancosMN, 2) = nTemp(2)
            matArregloBancosMN(nContBancosMN, 3) = nTemp(3)
            matArregloBancosMN(nContBancosMN, 4) = nTemp(4)
            matArregloBancosMN(nContBancosMN, 6) = nTemp(6)
            nTemp(1) = 0
            nTemp(2) = 0
            nTemp(3) = 0
            nTemp(4) = 0
            nTemp(6) = 0
        End If
    End If
       
    'Para Bancos en Moneda Nacional Saco los intereses
    Set rsTempo = GetInteresesCons("770154", dFechaAl, "0", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, "1")
    If rsTempo.EOF Then
    Else
        For nTem = 1 To nContBancosMN
            rsTempo.MoveFirst
            Do While Not rsTempo.EOF
                If Trim(UCase(matBancosMN(nTem))) = Trim(UCase(rsTempo!cBancoDesc)) Then
                    matArregloBancosMN(nTem, 5) = rsTempo!nPromInt
                End If
                rsTempo.MoveNext
            Loop
        Next
    End If
    Set rsTempo = Nothing
    
    'Cajas
    Do While UCase(xlHoja1.Cells(i, 1)) <> "CMACS"
        i = i + 1
    Loop
    
    i = i + 1
     
    nContCajasMN = 0
    nTemp(1) = 0
    nTemp(2) = 0
    nTemp(3) = 0
    nTemp(4) = 0
    nTemp(5) = 0
    nTemp(6) = 0
    
    'Obtengo los valores
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            If nContCajasMN > 0 Then
                matArregloCajasMN(nContCajasMN, 1) = nTemp(1)
                matArregloCajasMN(nContCajasMN, 2) = nTemp(2)
                matArregloCajasMN(nContCajasMN, 3) = nTemp(3)
                matArregloCajasMN(nContCajasMN, 5) = nTemp(5)
                nTemp(1) = 0
                nTemp(2) = 0
                nTemp(3) = 0
                nTemp(5) = 0
            End If
            nContCajasMN = nContCajasMN + 1
            matCajasMN(nContCajasMN) = Trim(xlHoja1.Cells(i, 1))
        End If
        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(i, 3)) ' Ahorro
        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(i, 5)) ' Plazo Fijo
        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(i, 7)) ' Total
        nTemp(5) = nTemp(5) + Val(xlHoja1.Cells(i, 8)) ' Endeudados
        i = i + 1
    Loop
    
    If nTemp(1) > 0 Or nTemp(2) > 0 Or nTemp(3) > 0 Or nTemp(5) > 0 Then
        If nContCajasMN > 0 Then
            matArregloCajasMN(nContCajasMN, 1) = nTemp(1)
            matArregloCajasMN(nContCajasMN, 2) = nTemp(2)
            matArregloCajasMN(nContCajasMN, 3) = nTemp(3)
            matArregloCajasMN(nContCajasMN, 5) = nTemp(5)
            nTemp(1) = 0
            nTemp(2) = 0
            nTemp(3) = 0
            nTemp(5) = 0
        End If
    End If
    
    'Para Cajas en Moneda Nacional Saco los intereses
    Set rsTempo = GetInteresesCons("770154", dFechaAl, "1", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, "1")
    If rsTempo.EOF Then
    Else
        For nTem = 1 To nContCajasMN
            rsTempo.MoveFirst
            Do While Not rsTempo.EOF
                If Trim(UCase(matCajasMN(nTem))) = Trim(UCase(rsTempo!cBancoDesc)) Then
                    matArregloCajasMN(nTem, 4) = rsTempo!nPromInt
                End If
                rsTempo.MoveNext
            Loop
        Next
    End If
    Set rsTempo = Nothing
    
    xlLibro.Close
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Abro el Anx15A_ConsolBancos_MMYYYYME.xls
     'BANCOS MONEDA EXTRANJERA
     
    lsArchivo = sArchConsolBcos(2)
    
    'Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(lsArchivo)

    bencontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If UCase(xlHoja1.Name) = UCase(Format(dFechaAl, "dd-mm-yyyy")) Then
            bencontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bencontrado = False Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        MsgBox "No existen datos con la fecha especificada" & Chr(13) & " en el archivo " & lsArchivo, vbExclamation, "Aviso!!!"
        Exit Sub
    End If

    nTemp(1) = 0
    nTemp(2) = 0
    nTemp(3) = 0
    nTemp(4) = 0
    nTemp(5) = 0
    nTemp(6) = 0
      
      
      
'   *************************************
    
    For i = 1 To 50
        matTempo(i) = ""
    Next
    ntempo = 0
    nTempo_2 = 0
    K = 0
    nFilaIniRestrin = 0
    nFilaFinRestrin = 0
 
    i = 9
    nContBancosME = 0
    Do While UCase(xlHoja1.Cells(i, 1)) <> "CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            
            If Trim(xlHoja1.Cells(i, 1)) = "BANCOS RESTRINGIDOS" Then
                nFilaIniRestrin = i + 1
            ElseIf Trim(xlHoja1.Cells(i, 1)) = "SUBTOTAL BCOS. RESTRINGIDOS" Then
                nFilaFinRestrin = i - 1
            End If
            
            nTempo_2 = 0
            For K = 1 To ntempo
                If Trim(matTempo(K)) = Trim(xlHoja1.Cells(i, 1)) Then
                    nTempo_2 = nTempo_2 + 1
                End If
            Next
            If nTempo_2 = 0 Then
                If Trim(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BANCOS" And Trim(xlHoja1.Cells(i, 1)) <> "BANCOS RESTRINGIDOS" And Trim(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BCOS. RESTRINGIDOS" Then
                    nContBancosME = nContBancosME + 1
                    ntempo = ntempo + 1
                    matTempo(ntempo) = Trim(xlHoja1.Cells(i, 1))
                End If
            End If
        End If
        i = i + 1
    Loop
'   *************************************
      
'    I = 9
'    nContBancosME = 0
'
'    'Saco cantidad de Bancos
'    Do While UCase(xlHoja1.Cells(I, 1)) <> "SUBTOTAL BANCOS"
'        If Len(Trim(xlHoja1.Cells(I, 1))) > 0 Then
'            nContBancosME = nContBancosME + 1
'        End If
'        I = I + 1
'    Loop
'
'    'Saco cantidad de Cajas
'    Do While UCase(xlHoja1.Cells(I, 1)) <> "CMACS"
'        I = I + 1
'    Loop
    
    i = i + 1
    
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
            nContCajasME = nContCajasME + 1
        End If
        i = i + 1
    Loop
    
    'Bancos
    ReDim matBancosME(nContBancosME) As String
    ReDim matArregloBancosME(nContBancosME, 6) As Currency
            
    'Cajas
    ReDim matCajasME(nContCajasME) As String
    ReDim matArregloCajasME(nContCajasME, 5) As Currency
    
    i = 9
    nContBancosME = 0
    

    
    'Obtengo los valores
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL BANCOS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
        
            If nContBancosME > 0 Then
                matArregloBancosME(nContBancosME, 1) = nTemp(1)
                matArregloBancosME(nContBancosME, 2) = nTemp(2)
                matArregloBancosME(nContBancosME, 3) = nTemp(3)
                matArregloBancosME(nContBancosME, 4) = nTemp(4)
                matArregloBancosME(nContBancosME, 6) = nTemp(6)
                nTemp(1) = 0
                nTemp(2) = 0
                nTemp(3) = 0
                nTemp(4) = 0
                nTemp(6) = 0
            End If
            
            nContBancosME = nContBancosME + 1
            matBancosME(nContBancosME) = Trim(xlHoja1.Cells(i, 1))
        
        End If
        
        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(i, 5)) ' Ahorro
        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(i, 3)) ' Cta Cte
        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(i, 7)) ' Plazo Fijo
        nTemp(4) = nTemp(4) + Val(xlHoja1.Cells(i, 8)) ' Total
        nTemp(6) = nTemp(6) + Val(xlHoja1.Cells(i, 9)) ' Endeudados
        
        If nFilaIniRestrin = nFilaFinRestrin And nFilaIniRestrin = 0 Then
        Else
            If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
                mTempo = ""
                For m = nFilaIniRestrin To nFilaFinRestrin
                    If Len(Trim(xlHoja1.Cells(m, 1))) > 0 Then
                        mTempo = Trim(xlHoja1.Cells(m, 1))
                    End If
                    If Trim(mTempo) = Trim(matBancosME(nContBancosME)) Then
                        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(m, 5)) ' Ahorro
                        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(m, 3)) ' Cta Cte
                        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(m, 7)) ' Plazo Fijo
                        nTemp(4) = nTemp(4) + Val(xlHoja1.Cells(m, 8)) ' Total
                        nTemp(6) = nTemp(6) + Val(xlHoja1.Cells(m, 9)) ' Endeudados
                    End If
                Next
            End If
        End If
        
        i = i + 1
    Loop
    
    If nTemp(1) > 0 Or nTemp(2) > 0 Or nTemp(3) > 0 Or nTemp(4) > 0 Or nTemp(6) > 0 Then
        If nContBancosME > 0 Then
            matArregloBancosME(nContBancosME, 1) = nTemp(1)
            matArregloBancosME(nContBancosME, 2) = nTemp(2)
            matArregloBancosME(nContBancosME, 3) = nTemp(3)
            matArregloBancosME(nContBancosME, 4) = nTemp(4)
            matArregloBancosME(nContBancosME, 6) = nTemp(6)
            nTemp(1) = 0
            nTemp(2) = 0
            nTemp(3) = 0
            nTemp(4) = 0
            nTemp(6) = 0
        End If
    End If
     
    'Para Bancos en Moneda Extranjera Saco los intereses
    Set rsTempo = GetInteresesCons("770154", dFechaAl, "0", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, "2")
    If rsTempo.EOF Then
    Else
        For nTem = 1 To nContBancosME
            rsTempo.MoveFirst
            Do While Not rsTempo.EOF
                If Trim(UCase(matBancosME(nTem))) = Trim(UCase(rsTempo!cBancoDesc)) Then
                    matArregloBancosME(nTem, 5) = rsTempo!nPromInt
                End If
                rsTempo.MoveNext
            Loop
        Next
    End If
    Set rsTempo = Nothing
     
     
    'Cajas
    Do While UCase(xlHoja1.Cells(i, 1)) <> "CMACS"
        i = i + 1
    Loop
    
    i = i + 1
    
    nContCajasME = 0
    nTemp(1) = 0
    nTemp(2) = 0
    nTemp(3) = 0
    nTemp(4) = 0
    nTemp(5) = 0
    nTemp(6) = 0
    
    'Obtengo los valores
    Do While UCase(xlHoja1.Cells(i, 1)) <> "SUBTOTAL CMACS"
        If Len(Trim(xlHoja1.Cells(i, 1))) > 0 Then
        
            If nContCajasME > 0 Then
                matArregloCajasME(nContCajasME, 1) = nTemp(1)
                matArregloCajasME(nContCajasME, 2) = nTemp(2)
                matArregloCajasME(nContCajasME, 3) = nTemp(3)
                matArregloCajasME(nContCajasME, 5) = nTemp(5)
                nTemp(1) = 0
                nTemp(2) = 0
                nTemp(3) = 0
                nTemp(5) = 0
            End If
            
            nContCajasME = nContCajasME + 1
            matCajasME(nContCajasME) = Trim(xlHoja1.Cells(i, 1))
        
        End If
        
        'Cambio 8x7 9x8
        nTemp(1) = nTemp(1) + Val(xlHoja1.Cells(i, 3)) ' Ahorro
        nTemp(2) = nTemp(2) + Val(xlHoja1.Cells(i, 5)) ' Plazo Fijo
        nTemp(3) = nTemp(3) + Val(xlHoja1.Cells(i, 7)) ' Total
        nTemp(5) = nTemp(5) + Val(xlHoja1.Cells(i, 8)) ' Endeudados
        i = i + 1
    Loop
    
    If nTemp(1) > 0 Or nTemp(2) > 0 Or nTemp(3) > 0 Or nTemp(5) > 0 Then
        If nContCajasME > 0 Then
            matArregloCajasME(nContCajasME, 1) = nTemp(1)
            matArregloCajasME(nContCajasME, 2) = nTemp(2)
            matArregloCajasME(nContCajasME, 3) = nTemp(3)
            matArregloCajasME(nContCajasME, 5) = nTemp(5)
            nTemp(1) = 0
            nTemp(2) = 0
            nTemp(3) = 0
            nTemp(5) = 0
        End If
    End If
     
    'Para Cajas en Moneda Extranjera Saco los intereses
    Set rsTempo = GetInteresesCons("770154", dFechaAl, "1", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, "2")
    If rsTempo.EOF Then
    Else
        For nTem = 1 To nContCajasME
            rsTempo.MoveFirst
            Do While Not rsTempo.EOF
                If Trim(UCase(matCajasME(nTem))) = Trim(UCase(rsTempo!cBancoDesc)) Then
                    matArregloCajasME(nTem, 4) = rsTempo!nPromInt
                End If
                rsTempo.MoveNext
            Loop
        Next
    End If
    Set rsTempo = Nothing

    xlLibro.Close

    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' BANCOS
    '''''''''''
    
    lbLibroOpen = ExcelBegin(sArchGrabar, xlAplicacion, xlLibro)
     
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       
       xlHoja1.Name = "BANCA MULTIPLE"
       xlHoja1.Cells(1, 2) = "SALDOS DE NUESTRAS CUENTAS EN BANCA MULTIPLE"
       xlHoja1.Cells(1, 8) = "'" & Format(dFechaAl, "dd/MM/YYYY")
       xlHoja1.Range("B1:H1").Font.Bold = True
       
       'MONEDA NACIONAL
       
       xlHoja1.Cells(4, 2) = "DEPOSITOS EN MONEDA NACIONAL"
'       nFil = nFil + 1
       xlHoja1.Cells(5, 3) = "AHORRO MN"
       xlHoja1.Cells(5, 4) = "CTE MN"
       xlHoja1.Cells(5, 5) = "PF MN"
       xlHoja1.Cells(5, 6) = "TOTAL"
       xlHoja1.Cells(5, 7) = "TEA PF"
       xlHoja1.Cells(5, 8) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B4:H5").Font.Bold = True
       
       nfil = 5
       
       For i = 1 To nContBancosMN
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = matBancosMN(i)
            For j = 1 To 6
                
                xlHoja1.Cells(nfil, j + 2) = matArregloBancosMN(i, j)
                
            Next
       Next
       
       ''''''''''''''''''''''
       For i = 1 To nContBancosME
            nTemp1 = 0
            For j = 1 To nContBancosMN
                If matBancosME(i) = matBancosMN(j) Then
                    nTemp1 = nTemp1 + 1
                End If
            Next
            If nTemp1 = 0 Then
                nfil = nfil + 1
                xlHoja1.Cells(nfil, 2) = matBancosME(i)
                xlHoja1.Cells(nfil, 3) = 0
                xlHoja1.Cells(nfil, 4) = 0
                xlHoja1.Cells(nfil, 5) = 0
                xlHoja1.Cells(nfil, 6) = 0
                xlHoja1.Cells(nfil, 7) = 0
                xlHoja1.Cells(nfil, 8) = 0
            End If
        Next
       
       ''''''''''''''''''''''''''''''''''''''
       'ORDENO
       xlHoja1.Range("B6:H" & Trim(Str(nfil))).Sort Key1:=xlHoja1.Range("B5")
       ''''''''''''''''''''''''''''''''''''''
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C6:C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D6:D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E6:E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F6:F" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=SUM(G6:G" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=SUM(H6:H" & Trim(Str(nfil - 1)) & ")"
       
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, 5, 8, 5
       ExcelCuadro xlHoja1, 2, 6, 8, nfil - 1
       ExcelCuadro xlHoja1, 2, nfil, 8, nfil
       
       'MONEDA EXTRANJERA
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "DEPOSITOS EN MONEDA EXTRANJERA"
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 3) = "AHORRO ME"
       xlHoja1.Cells(nfil, 4) = "CTE ME"
       xlHoja1.Cells(nfil, 5) = "PF ME"
       xlHoja1.Cells(nfil, 6) = "TOTAL"
       xlHoja1.Cells(nfil, 7) = "TEA PF"
       xlHoja1.Cells(nfil, 8) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).Font.Bold = True

       nFilTemp1 = nfil + 1
       
       For i = 1 To nContBancosME
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = matBancosME(i)
            For j = 1 To 6
                xlHoja1.Cells(nfil, j + 2) = matArregloBancosME(i, j)
            Next
       Next
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''
       For i = 1 To nContBancosMN
            nTemp1 = 0
            For j = 1 To nContBancosME
                If matBancosMN(i) = matBancosME(j) Then
                    nTemp1 = nTemp1 + 1
                End If
            Next
            If nTemp1 = 0 Then
                nfil = nfil + 1
                xlHoja1.Cells(nfil, 2) = matBancosMN(i)
                xlHoja1.Cells(nfil, 3) = 0
                xlHoja1.Cells(nfil, 4) = 0
                xlHoja1.Cells(nfil, 5) = 0
                xlHoja1.Cells(nfil, 6) = 0
                xlHoja1.Cells(nfil, 7) = 0
                xlHoja1.Cells(nfil, 8) = 0
            End If
       Next
       
       '''''''''''''''''''''''''''''''''''''''''''''''''''
       'ORDENO
       xlHoja1.Range("B" & Trim(Str(nFilTemp1)) & ":H" & Trim(Str(nfil))).Sort Key1:=xlHoja1.Range("B" & Trim(Str(nFilTemp1)))
       ''''''''''''''''''''''''''''''''''''''''''''''''''
       nFilTemp2 = nfil
       
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nFilTemp1)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nFilTemp1)) & ":D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E" & Trim(Str(nFilTemp1)) & ":E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F" & Trim(Str(nFilTemp1)) & ":F" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=SUM(G" & Trim(Str(nFilTemp1)) & ":G" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=SUM(H" & Trim(Str(nFilTemp1)) & ":H" & Trim(Str(nfil - 1)) & ")"
       
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, nFilTemp1 - 1, 8, nFilTemp1 - 1
       ExcelCuadro xlHoja1, 2, nFilTemp1, 8, nfil
       ExcelCuadro xlHoja1, 2, nfil, 8, nfil
       
       
       'FIN MONEDA NACIONAL Y EXTRANJERA
       
       'CONSOLIDADO A TC
       
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "CONSOLIDADO A T.C."
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 3) = "AHORRO"
       xlHoja1.Cells(nfil, 4) = "CORRIENTE"
       xlHoja1.Cells(nfil, 5) = "PLAZO FIJO"
       xlHoja1.Cells(nfil, 6) = "TOTAL"
       xlHoja1.Cells(nfil, 7) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       For i = nFilTemp1 To nFilTemp2
           nfil = nfil + 1
           xlHoja1.Cells(nfil, 2) = xlHoja1.Cells(i, 2)
           xlHoja1.Cells(nfil, 3) = (Val(xlHoja1.Cells(i, 3)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 3))
           xlHoja1.Cells(nfil, 4) = (Val(xlHoja1.Cells(i, 4)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 4))
           xlHoja1.Cells(nfil, 5) = (Val(xlHoja1.Cells(i, 5)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 5))
           xlHoja1.Cells(nfil, 6) = (Val(xlHoja1.Cells(i, 6)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 6))
           xlHoja1.Cells(nfil, 7) = (Val(xlHoja1.Cells(i, 8)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 8))
       Next
            
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nFilTemp2 + 4)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nFilTemp2 + 4)) & ":D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E" & Trim(Str(nFilTemp2 + 4)) & ":E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F" & Trim(Str(nFilTemp2 + 4)) & ":F" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=SUM(G" & Trim(Str(nFilTemp2 + 4)) & ":G" & Trim(Str(nfil - 1)) & ")"
        
       nTemp2 = xlHoja1.Cells(nfil, 6)
       
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, nFilTemp2 + 3, 7, nFilTemp2 + 3
       ExcelCuadro xlHoja1, 2, nFilTemp2 + 4, 7, nfil
       ExcelCuadro xlHoja1, 2, nfil, 7, nfil
        
       'PARTICIPACION PORCENTUAL
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "PARTICIPACION PORCENTUAL"
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "BANCO"
       xlHoja1.Cells(nfil, 3) = "TOTAL S/."
       xlHoja1.Cells(nfil, 4) = "%"
       xlHoja1.Cells(nfil, 5) = "CLASIFICACION"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).Font.Bold = True
       
       xlHoja1.Cells(nfil + 1, 6) = "Liq. S/."
       xlHoja1.Cells(nfil + 1, 7) = nLiquid(1)
       xlHoja1.Cells(nfil + 2, 6) = "Liq. US $"
       xlHoja1.Cells(nfil + 2, 7) = nLiquid(2)
       
       nTemp1 = nfil + 1
       For i = nfil To nfil + (nFilTemp2 - nFilTemp1)
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = xlHoja1.Cells(nfil - (4 + (nFilTemp2 - nFilTemp1)), 2)
            xlHoja1.Cells(nfil, 3) = xlHoja1.Cells(nfil - (4 + (nFilTemp2 - nFilTemp1)), 6)
            
            If nPatrEfectivo = 0 Then
                xlHoja1.Cells(nfil, 4) = 0
            Else
                xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=+C" & Trim(Str(nfil)) & "*100/$C$" & Trim(Str(nTemp1 + (nFilTemp2 - nFilTemp1) + 4))
            End If
       
       Next
       
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "SUBTOTAL"
       xlHoja1.Cells(nfil, 3) = nTemp2
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nTemp1)) & ":D" & Trim(Str(nfil - 1)) & ")"
            
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil + 7))).Font.Bold = True
       
       
       ExcelCuadro xlHoja1, 2, nTemp1 - 1, 5, nTemp1 - 1
       ExcelCuadro xlHoja1, 2, nTemp1, 5, nfil - 1
       ExcelCuadro xlHoja1, 2, nfil, 5, nfil
            
       ExcelCuadro xlHoja1, 6, nTemp1, 7, nTemp1 + 1
       
       ExcelCuadro xlHoja1, 2, nfil + 1, 5, nfil + 1
       ExcelCuadro xlHoja1, 2, nfil + 2, 5, nfil + 2
       ExcelCuadro xlHoja1, 2, nfil + 3, 4, nfil + 3
       ExcelCuadro xlHoja1, 2, nfil + 4, 3, nfil + 4
       ExcelCuadro xlHoja1, 2, nfil + 5, 3, nfil + 5
       
       ExcelCuadro xlHoja1, 2, nfil + 7, 3, nfil + 7
        
       nValTemp2(1) = "='BANCA MULTIPLE'!C" & Trim(Str(nfil))
       
       nValTemp3(1) = "='BANCA MULTIPLE'!D" & Trim(Str(nfil))
            
 
       nfil = nfil + 1
        
       nValTemp1(1) = nfil 'para la cmac
       
       xlHoja1.Cells(nfil, 2) = "CMACS"
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "TOTAL"
       xlHoja1.Cells(nfil, 6) = "US $"
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nfil - 2)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Cells(nfil, 5) = Format(Val(xlHoja1.Cells(nfil, 3)) / nTipoCambio, "0.00")
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "PATRIM EFE"
       xlHoja1.Cells(nfil, 3) = nPatrEfectivo
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nfil - 3)) & ":D" & Trim(Str(nfil - 2)) & ")"
       
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "MAXIMO"
       xlHoja1.Cells(nfil, 3) = nPatrEfectivo * 0.3
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "1 o/o"
       xlHoja1.Cells(nfil, 3) = nPatrEfectivo * 0.01
       
       nfil = nfil + 2
       xlHoja1.Cells(nfil, 2) = "TIPO CAMBIO"
       
       nTempTC(1) = nfil
       
''''       If nPatrEfectivo = 0 Then
''''            nFil = nFil + 3
''''            xlHoja1.Cells(nFil, 2) = "Nota: Falta generacion Reporte 3 Para Obtener Patrim. EFE"
''''            xlHoja1.Cells(nFil, 2).Font.ColorIndex = 3
''''       End If
        
       xlHoja1.Cells.NumberFormat = "##,###,##0.00"
       xlHoja1.Cells.Font.Name = "Arial"
       xlHoja1.Cells.Font.Size = 8
       xlHoja1.Cells.EntireColumn.AutoFit
       
       '''''''''
       ' CAJAS
       '''''''''

       nTemp1 = 0
       nTemp2 = 0
       
       Set xlHoja1 = xlLibro.Worksheets(2)
       xlHoja1.Name = "SISTEMA DE CAJAS"
       xlHoja1.Cells(1, 2) = "SALDOS DE NUESTRAS CUENTAS EN CAJAS MUNICIPALES"
       xlHoja1.Cells(1, 7) = "'" & Format(dFechaAl, "dd/MM/YYYY")
       
       xlHoja1.Range("B1:H1").Font.Bold = True
       
       'MONEDA NACIONAL
       
       xlHoja1.Cells(4, 2) = "DEPOSITOS EN MONEDA NACIONAL"
'       nFil = nFil + 1
       xlHoja1.Cells(5, 3) = "AHORRO MN"
       xlHoja1.Cells(5, 4) = "PF MN"
       xlHoja1.Cells(5, 5) = "TOTAL"
       xlHoja1.Cells(5, 6) = "TEA PF"
       xlHoja1.Cells(5, 7) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B4:H5").Font.Bold = True
       
       nfil = 5
       
       For i = 1 To nContCajasMN
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = matCajasMN(i)
            For j = 1 To 5
                xlHoja1.Cells(nfil, j + 2) = matArregloCajasMN(i, j)
            Next
       Next
       
       For i = 1 To nContCajasME
            nTemp1 = 0
            For j = 1 To nContCajasMN
                If matCajasME(i) = matCajasMN(j) Then
                    nTemp1 = nTemp1 + 1
                End If
            Next
            If nTemp1 = 0 Then
                nfil = nfil + 1
                xlHoja1.Cells(nfil, 2) = matCajasME(i)
                xlHoja1.Cells(nfil, 3) = 0
                xlHoja1.Cells(nfil, 4) = 0
                xlHoja1.Cells(nfil, 5) = 0
                xlHoja1.Cells(nfil, 6) = 0
                xlHoja1.Cells(nfil, 7) = 0
            End If
        Next
       
       ''''''''''''''''''''''''''''''''''''''
       'ORDENO
       xlHoja1.Range("B6:G" & Trim(Str(nfil))).Sort Key1:=xlHoja1.Range("B5")
       ''''''''''''''''''''''''''''''''''''''
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C6:C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D6:D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E6:E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F6:F" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=SUM(G6:G" & Trim(Str(nfil - 1)) & ")"
       
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, 5, 7, 5
       ExcelCuadro xlHoja1, 2, 6, 7, nfil - 1
       ExcelCuadro xlHoja1, 2, nfil, 7, nfil
       
       'MONEDA EXTRANJERA
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "DEPOSITOS EN MONEDA EXTRANJERA"
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 3) = "AHORRO ME"
       xlHoja1.Cells(nfil, 4) = "PF ME"
       xlHoja1.Cells(nfil, 5) = "TOTAL"
       xlHoja1.Cells(nfil, 6) = "TEA PF"
       xlHoja1.Cells(nfil, 7) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       nFilTemp1 = nfil + 1
       
       For i = 1 To nContCajasME
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = matCajasME(i)
            For j = 1 To 5
                xlHoja1.Cells(nfil, j + 2) = matArregloCajasME(i, j)
            Next
       Next
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''
       For i = 1 To nContCajasMN
            nTemp1 = 0
            For j = 1 To nContCajasME
                If matCajasMN(i) = matCajasME(j) Then
                    nTemp1 = nTemp1 + 1
                End If
            Next
            If nTemp1 = 0 Then
                nfil = nfil + 1
                xlHoja1.Cells(nfil, 2) = matCajasMN(i)
                xlHoja1.Cells(nfil, 3) = 0
                xlHoja1.Cells(nfil, 4) = 0
                xlHoja1.Cells(nfil, 5) = 0
                xlHoja1.Cells(nfil, 6) = 0
                xlHoja1.Cells(nfil, 7) = 0
            End If
       Next
       
       '''''''''''''''''''''''''''''''''''''''''''''''''''
       'ORDENO
       xlHoja1.Range("B" & Trim(Str(nFilTemp1)) & ":G" & Trim(Str(nfil))).Sort Key1:=xlHoja1.Range("B" & Trim(Str(nFilTemp1)))
         
       ''''''''''''''''''''''''''''''''''''''''''''''''''
       nFilTemp2 = nfil
       
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nFilTemp1)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nFilTemp1)) & ":D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E" & Trim(Str(nFilTemp1)) & ":E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F" & Trim(Str(nFilTemp1)) & ":F" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=SUM(G" & Trim(Str(nFilTemp1)) & ":G" & Trim(Str(nfil - 1)) & ")"
        
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, nFilTemp1 - 1, 7, nFilTemp1 - 1
       ExcelCuadro xlHoja1, 2, nFilTemp1, 7, nfil
       ExcelCuadro xlHoja1, 2, nfil, 7, nfil
        
       'FIN MONEDA NACIONAL Y EXTRANJERA
       
       'CONSOLIDADO A TC
       
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "CONSOLIDADO A T.C."
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 3) = "AHORRO"
       xlHoja1.Cells(nfil, 4) = "PLAZO FIJO"
       xlHoja1.Cells(nfil, 5) = "TOTAL"
       xlHoja1.Cells(nfil, 6) = "ENDEUDAMIENTO"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       For i = nFilTemp1 To nFilTemp2
           nfil = nfil + 1
           xlHoja1.Cells(nfil, 2) = xlHoja1.Cells(i, 2)
           xlHoja1.Cells(nfil, 3) = (Val(xlHoja1.Cells(i, 3)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 3))
           xlHoja1.Cells(nfil, 4) = (Val(xlHoja1.Cells(i, 4)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 4))
           xlHoja1.Cells(nfil, 5) = (Val(xlHoja1.Cells(i, 5)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 5))
           xlHoja1.Cells(nfil, 6) = (Val(xlHoja1.Cells(i, 7)) * nTipoCambio) + Val(xlHoja1.Cells(i - (4 + (nFilTemp2 - nFilTemp1)), 7))
       Next
       
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "TOTAL"
        
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nFilTemp2 + 4)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nFilTemp2 + 4)) & ":D" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=SUM(E" & Trim(Str(nFilTemp2 + 4)) & ":E" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=SUM(F" & Trim(Str(nFilTemp2 + 4)) & ":F" & Trim(Str(nfil - 1)) & ")"
                
       nTemp2 = xlHoja1.Cells(nfil, 5)
       
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, nFilTemp2 + 3, 6, nFilTemp2 + 3
       ExcelCuadro xlHoja1, 2, nFilTemp2 + 4, 6, nfil
       ExcelCuadro xlHoja1, 2, nfil, 6, nfil
       
       '''''''''''''
       'PARTICIPACION PORCENTUAL
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "PARTICIPACION PORCENTUAL"
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "BANCO"
       xlHoja1.Cells(nfil, 3) = "TOTAL S/."
       xlHoja1.Cells(nfil, 4) = "%"
       xlHoja1.Cells(nfil, 5) = "CLASIFICACION"
       
       xlHoja1.Range("B" & Trim(Str(nfil - 1)) & ":G" & Trim(Str(nfil))).Font.Bold = True
       
       nTemp1 = nfil + 1
       For i = nfil To nfil + (nFilTemp2 - nFilTemp1)
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = xlHoja1.Cells(nfil - (4 + (nFilTemp2 - nFilTemp1)), 2)
            xlHoja1.Cells(nfil, 3) = xlHoja1.Cells(nfil - (4 + (nFilTemp2 - nFilTemp1)), 5)
            If nPatrEfectivo = 0 Then
                xlHoja1.Cells(nfil, 4) = 0
            Else
                'xlHoja1.Cells(nFil, 4) = Format(Val(xlHoja1.Cells(nFil - (4 + (nFilTemp2 - nFilTemp1)), 5)) / nPatrEfectivo * 100, "0.00")
                xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=+C" & Trim(Str(nfil)) & "*100/$C$" & Trim(Str(nTemp1 + (nFilTemp2 - nFilTemp1) + 4))
            End If
            
       Next
       
       nfil = nfil + 1
       
       xlHoja1.Cells(nfil, 2) = "SUBTOTAL"
       xlHoja1.Cells(nfil, 3) = nTemp2
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=SUM(D" & Trim(Str(nTemp1)) & ":D" & Trim(Str(nfil - 1)) & ")"
            
       xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil + 7))).Font.Bold = True
       
       ExcelCuadro xlHoja1, 2, nTemp1 - 1, 5, nTemp1 - 1
       ExcelCuadro xlHoja1, 2, nTemp1, 5, nfil - 1
       ExcelCuadro xlHoja1, 2, nfil, 5, nfil
       
       ExcelCuadro xlHoja1, 2, nfil + 1, 5, nfil + 1
       ExcelCuadro xlHoja1, 2, nfil + 2, 5, nfil + 2
       ExcelCuadro xlHoja1, 2, nfil + 3, 4, nfil + 3
       ExcelCuadro xlHoja1, 2, nfil + 5, 3, nfil + 5
 
       nValTemp2(2) = "='SISTEMA DE CAJAS'!C" & Trim(Str(nfil))
 
       nValTemp3(2) = "='SISTEMA DE CAJAS'!D" & Trim(Str(nfil))
       
       nfil = nfil + 1
       nValTemp1(2) = nfil 'para la cmac
       
       xlHoja1.Cells(nfil, 2) = "BANCOS"
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = nValTemp2(1)
       
       xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = nValTemp3(1)
       
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "TOTAL"
       xlHoja1.Cells(nfil, 6) = "US $"
       xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=SUM(C" & Trim(Str(nfil - 2)) & ":C" & Trim(Str(nfil - 1)) & ")"
       xlHoja1.Cells(nfil, 5) = Format(Val(xlHoja1.Cells(nfil, 3)) / nTipoCambio, "0.00")
       nfil = nfil + 1
       xlHoja1.Cells(nfil, 2) = "PATRIM EFE"
       xlHoja1.Cells(nfil, 3) = nPatrEfectivo
       
       nfil = nfil + 2
       xlHoja1.Cells(nfil, 2) = "TIPO CAMBIO"
       
       
       nTempTC(2) = nfil
       
''''       If nPatrEfectivo = 0 Then
''''            nFil = nFil + 3
''''            xlHoja1.Cells(nFil, 2) = "Nota: Falta generacion Reporte 3 Para Obtener Patrim. EFE"
''''            xlHoja1.Cells(nFil, 2).Font.ColorIndex = 3
''''       End If
       
        '''''''''''''''''
        'ACTUALIZO
       Set xlHoja1 = xlLibro.Worksheets(1)
       xlHoja1.Range("C" & Trim(Str(nValTemp1(1))) & ":C" & Trim(Str(nValTemp1(1)))).Formula = nValTemp2(2)
       
       xlHoja1.Range("D" & Trim(Str(nValTemp1(1))) & ":D" & Trim(Str(nValTemp1(1)))).Formula = nValTemp3(2)
       
       Set xlHoja1 = xlLibro.Worksheets(2)
       
       xlHoja1.Cells.NumberFormat = "##,###,##0.00"
       xlHoja1.Cells.Font.Name = "Arial"
       xlHoja1.Cells.Font.Size = 8
       xlHoja1.Cells.EntireColumn.AutoFit
       
       
       xlLibro.Sheets("BANCA MULTIPLE").Cells(nTempTC(1), 3) = nTipoCambio
       xlLibro.Sheets("BANCA MULTIPLE").Range("C" & Trim(Str(nTempTC(1)))).NumberFormat = "0.###0"
        
       xlLibro.Sheets("SISTEMA DE CAJAS").Cells(nTempTC(2), 3) = nTipoCambio
       xlLibro.Sheets("SISTEMA DE CAJAS").Range("C" & Trim(Str(nTempTC(2)))).NumberFormat = "0.###0"
       
       ExcelEnd sArchGrabar, xlAplicacion, xlLibro, xlHoja1
       lbLibroOpen = False
       CargaArchivo "Concfondos" & Format(dFechaAl, "ddMMYYYY") & ".XLS", App.path & "\SPOOLER"
    End If
     
 
 
Exit Sub
ErrBegin:
  
  ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
 
    
End Sub

Private Function GetInteresesCons(ByVal psOpeCod As String, ByVal pdFecha As Date, ByVal psOpeObjOrden As String, psCtaIFEstado As String, psMoneda As String) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

'gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida

sql = "SELECT  CI.cPersCod, CtaIF.cBancoDesc, AVG(CII.nCtaIFIntValor)  nPromInt " & _
         " FROM    CTAIFSALDO CI " & _
         " LEFT JOIN (Select cPersCod, cIFTpo, cCtaIFCod, nCtaIFIntValor FROM CtaIFInteres CII2 " & _
         " WHERE CII2.dCtaIFIntRegistro = (Select Max(dCtaIFIntRegistro) FROM CtaIFInteres CII1 WHERE CII1.cPErsCod = CII2.cPersCod and CII1.cIFTpo = CII2.cIFTpo and CII1.cCtaIFCod = CII2.cCtaIFCod) " & _
         " ) CII ON CII.cPErsCod = CI.cPersCod and CII.cIFTpo = CI.cIFTpo and CII.cCtaIFCod = CI.cCtaIFCod " & _
         " Join " & _
         " ( SELECT  CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod AS cCtaIfCod, CI.cCtaIFDesc, P.cPersNombre cBancoDesc, CI.cCtaIFEstado , CI.nInteres " & _
         " FROM    CTAIF CI JOIN INSTITUCIONFINANC I ON  I.CPERSCOD = CI.CPERSCOD  AND CI.CIFTPO = I.CIFTPO " & _
         " JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD, OPEOBJ O " & _
         " WHERE   O.cOpeCod = '" & Trim(psOpeCod) & "' " & _
         " and O.cOpeObjOrden IN ('" & Trim(psOpeObjOrden) & "') " & _
         " and CI.cCtaIFEstado IN ('" & Trim(psCtaIFEstado) & "') " & _
         " AND CI.cIFTpo + CI.cCtaIFCod like O.cOpeObjFiltro " & _
         " ) as CtaIf " & _
         " ON  SUBSTRING(CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod , 1, LEN(CtaIf.cCtaIfCod)) = CtaIf.cCtaIfCod , " & _
         " OPEOBJ O JOIN OpeCta oc ON oc.cOpeCod = o.cOpeCod " & _
         " WHERE   O.cOpeCod = '" & Trim(psOpeCod) & "' " & _
         " and ci.cCtaContCod LIKE LEFT(oc.cCtaContCod,2)+'" & Trim(psMoneda) & "'+SubString(oc.cCtaContCod,4,22) + '%' and oc.cOpeCtaOrden IN ('" & Trim(psOpeObjOrden) & "') " & _
         " and O.cOpeObjOrden IN ('" & Trim(psOpeObjOrden) & "') " & _
         " AND CI.cIFTpo + CI.cCtaIFCod Like O.cOpeObjFiltro  and CI.cCtaIFCod LIKE '03" & Trim(Str(psMoneda)) & "%'" & _
         " AND CONVERT(varchar(8), CI.dCtaIFSaldo,112)<='" & Format(pdFecha, "YYYYMMdd") & "' " & _
         " GROUP   BY CI.cPersCod, CtaIF.cBancoDesc"
            
Set rs = oConec.CargaRecordSet(sql)
Set GetInteresesCons = rs
oConec.CierraConexion
Set oConec = Nothing
            
End Function


Public Sub ResumenChqDepositados(pdFecha As Date, psOpeCod As String)
    Dim lFecha As String
    Dim lnTotal As Currency
    Dim sTexto  As String
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim oCon As DConecta
    Dim oOpe As DOperacion
    Dim lnLin As Integer
    Dim lnPag As Integer
    Set oCon = New DConecta
    oCon.AbreConexion

    lFecha = pdFecha
    lnTotal = 0
    lnPag = 0
    lnLin = gnLinPage
    sTexto = ""

    sSql = "SELECT  mDep.nMovNro, mDep.cMovNro, RTRIM(p.cPersNombre) as Banco, RTRIM(pCh.cPersNombre) as BancoCh, " _
         & "        dr.cNroDoc, md.dDocFecha as Registro, " _
         & "        mc.nMovImporte, dr.nMonto nMontoChq, dr.cIFCta cCtaBco, ISNULL(drc.cCtaCod,'') cCtaCod, ISNULL(ag.cAgeDescripcion,a.cAreaDescripcion) Agencia " _
         & "FROM   Mov mDep " _
         & "       JOIN MovCta mc ON mc.nMovNro = mDep.nMovNro " _
         & "       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
         & "       JOIN Persona p ON p.cPersCod = mif.cPersCod " _
         & "       JOIN CtaIF ci ON ci.cPersCod = mif.cPErsCod and ci.cIFTpo = mif.cIFTpo and ci.cCtaIFCod = mif.cCtaIFCod " _
         & "       JOIN MovRef mr ON mr.nMovNro = mDep.nMovNro " _
         & "       JOIN Mov m ON m.nMovNro = mr.nMovNroRef " _
         & "       JOIN MovDoc md on md.nMovNro = m.nMovNro " _
         & "       JOIN DocRec dr ON dr.cNroDoc = md.cDocNro and dr.nTpoDoc = md.nDocTpo " _
         & "       JOIN DocRecEst dre ON dre.cNroDoc = dr.cNroDoc and dre.nTpoDoc = dr.nTpoDoc and dre.cMovNro = m.cMovNro JOIN Persona pCh ON pCh.cPersCod = dr.cPersCod " _
         & "       LEFT JOIN DocRecCapta drc ON dr.cPersCod = drc.cPersCod and dr.nTpoDoc = drc.nTpoDoc and dr.cNroDoc = drc.cNroDoc " _
         & "       LEFT JOIN Areas a ON a.cAreaCod = dr.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = dr.cAgeCod " _
         & "WHERE   dr.nMoneda = '" & Mid(psOpeCod, 3, 1) & "' and mDep.nMovFlag not in (1,5) and m.nMovFlag not in (1,5) " _
         & "        and left(mDep.cmovnro,8) = '" & Format(pdFecha, gsFormatoMovFecha) & "' " _
         & "Order by mDep.nMovNro "


    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        rs.Close
        MsgBox "No se depositaron Cheques en la Fecha", vbInformation, "Error"
        Exit Sub
    End If
    lnTotal = 0
    sTexto = PrnSet("C+") + PrnSet("B+")
    Do While Not rs.EOF
       If lnLin > gnLinPage - 4 Then
          Linea sTexto, CabeRepo(gsNomCmac, gsNomAge, "Finanzas", "", Format(gdFecSis, gsFormatoFechaView), "LISTADO  DE CHEQUES DEPOSITADOS" & " EN " & IIf(Mid(psOpeCod, 3, 1) = "1", "M.N.", "M.E."), " Fecha : " & lFecha, "", "", lnPag, gnColPage), 0, lnLin
          lnLin = 6
          Linea sTexto, String(145, "="), , lnLin
          Linea sTexto, ImpreFormat("NRO. MOVIMIENTO", 25) & ImpreFormat("BANCO DEPOSITO", 30) & ImpreFormat("MONTO DEPOSITO", 15), , lnLin
          Linea sTexto, String(145, "-") + PrnSet("B-"), , lnLin
       End If
       lnTotal = lnTotal + rs!nMovImporte
       gsMovNro = rs!cMovNro
       Linea sTexto, "", , lnLin
       Linea sTexto, ImpreFormat(rs!cMovNro, 25) & ImpreFormat(rs!banco, 30) & ImpreFormat(rs!nMovImporte, 12, , True), , lnLin
       Linea sTexto, ImpreFormat("NRO CHEQUE", 14) & ImpreFormat("FEC.RECEP.", 12, 0) & ImpreFormat("BANCO", 25) & ImpreFormat("CTA.BANCO", 14) & ImpreFormat("AGENCIA", 25) & ImpreFormat("CTA CMAC", 21) & ImpreFormat("MONTO", 8), , lnLin
       Do While gsMovNro = rs!cMovNro
          If lnLin > gnLinPage - 4 Then
              Linea sTexto, CabeRepo(gsNomCmac, gsNomAge, "Finanzas", "", Format(gdFecSis, gsFormatoFechaView), "LISTADO  DE CHEQUES DEPOSITADOS" & " EN " & IIf(Mid(psOpeCod, 3, 1) = "1", "M.N.", "M.E."), " Fecha : " & lFecha, "", "", lnPag, gnColPage), 0, lnLin
              lnLin = 6
              Linea sTexto, String(145, "="), , lnLin
              Linea sTexto, ImpreFormat("NRO. MOVIMIENTO", 25) & ImpreFormat("BANCO DEPOSITO", 30) & ImpreFormat("MONTO DEPOSITO", 15), , lnLin
              Linea sTexto, String(145, "-") + PrnSet("B-"), , lnLin
          End If
          Linea sTexto, ImpreFormat(rs!cNroDoc, 14) & ImpreFormat(rs!Registro, 12, 0) & _
              ImpreFormat(rs!BancoCh, 25) & ImpreFormat(rs!cCtaBco, 14) & ImpreFormat(rs!Agencia, 25) & _
              ImpreFormat(rs!cCtaCod, 15) & ImpreFormat(rs!nMontoChq, 13, 2, True), , lnLin
          rs.MoveNext
          If rs.EOF Then
            Exit Do
          End If
       Loop
    Loop
    RSClose rs
    sTexto = sTexto & oImpresora.gPrnSaltoLinea
    sTexto = sTexto + PrnSet("B+") + ImpreFormat("", 50) + ImpreFormat("TOTAL DEPOSITADO : " & gsSimbolo, 17) & ImpreFormat(lnTotal, 12, 2, True) & oImpresora.gPrnSaltoLinea & PrnSet("C-")
    EnviaPrevio sTexto, "Resumen de Cheques Recibidos", gnLinPage, False
End Sub

Public Sub ReporteAdeudadosVinculados(ByVal sOpeCod As String, Optional ByVal nTPEuros As Currency)
Dim oCtaIf As NCajaCtaIF
Dim lsMoneda As String
Dim rs As ADODB.Recordset

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String


If sOpeCod = OpeCGAdeudRepVinculadosMN Then
    lsMoneda = "1"
Else
   lsMoneda = "2"
End If
On Error GoTo ReporteAdeudadosVinculadosErr

    Set rs = New ADODB.Recordset
           
    Set oCtaIf = New NCajaCtaIF
    Set rs = oCtaIf.GetReporteAdeudadoVin(lsMoneda)   '--gnTipoCambioEuro)
    
    If rs Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    
    glsArchivo = "Reporte_Adeudados_Vinculados" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
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
            lsNomHoja = "Adeudos Vinculados"
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

            xlAplicacion.Range("A1:A1").ColumnWidth = 7
            xlAplicacion.Range("B1:B1").ColumnWidth = 37
            xlAplicacion.Range("c1:c1").ColumnWidth = 13
            xlAplicacion.Range("D1:D1").ColumnWidth = 20
            xlAplicacion.Range("E1:E1").ColumnWidth = 20
            xlAplicacion.Range("F1:F1").ColumnWidth = 10
            xlAplicacion.Range("G1:G1").ColumnWidth = 10
            xlAplicacion.Range("H1:H1").ColumnWidth = 13
            xlAplicacion.Range("I1:I1").ColumnWidth = 30
            xlAplicacion.Range("J1:J1").ColumnWidth = 10
            xlAplicacion.Range("K1:K1").ColumnWidth = 10
            xlAplicacion.Range("L1:L1").ColumnWidth = 40
           
            xlAplicacion.Range("A1:Z100").Font.Size = 9
            xlAplicacion.Range("A1:Z100").Font.Name = "Century Gothic"
       
            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 1) = "Detalle de Operaciones Back to Back Vigentes al " & Format(gdFecSis, "dd/mm/yyyy")
            
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Merge True
                 
                      
            liLineas = 4
            
            xlHoja1.Cells(liLineas, 1) = "N°"
            xlHoja1.Cells(liLineas, 2) = "Entidad"
            xlHoja1.Cells(liLineas, 3) = "Linea Original"
            xlHoja1.Cells(liLineas, 4) = "Recursos Recibidos y Depositados en garantia $"
            xlHoja1.Cells(liLineas, 5) = "Recursos Recibidos y Depositados en gaantia EUROS"
            xlHoja1.Cells(liLineas, 6) = "Plazo"
            xlHoja1.Cells(liLineas, 7) = "TEA"
            xlHoja1.Cells(liLineas, 8) = "Monto de Créditos obtenidos (Adeudado)S/."
            xlHoja1.Cells(liLineas, 9) = "Empresa del Sistema Financiero Nacional"
            xlHoja1.Cells(liLineas, 10) = "Plazo"
            xlHoja1.Cells(liLineas, 11) = "TEA"
            xlHoja1.Cells(liLineas, 12) = "Diferencial de la Operacion"
            
   
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Font.Bold = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Interior.ColorIndex = 36 '.Color = RGB(159, 206, 238)
            
         
            
            liLineas = liLineas + 1
                 
         Do Until rs.EOF
            xlHoja1.Cells(liLineas, 2) = rs(0)
            xlHoja1.Cells(liLineas, 3) = rs(1)
            xlHoja1.Cells(liLineas, 4) = rs(2)
            xlHoja1.Cells(liLineas, 5) = rs(3)
            xlHoja1.Cells(liLineas, 6) = rs(4) & " días"
            xlHoja1.Cells(liLineas, 7) = rs(5)
            xlHoja1.Cells(liLineas, 8) = rs(6)
            xlHoja1.Cells(liLineas, 9) = rs(7)
            xlHoja1.Cells(liLineas, 10) = rs(8) & " días"
            xlHoja1.Cells(liLineas, 11) = rs(9)
            
            
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).Style = "Comma"
            xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Style = "Comma"
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 11)).HorizontalAlignment = xlCenter
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 6), xlHoja1.Cells(liLineas, 6)).HorizontalAlignment = xlRight
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlRight
            
            liLineas = liLineas + 1
            rs.MoveNext
        Loop

        ExcelCuadro xlHoja1, 1, 4, 12, liLineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
  
    
Set oCtaIf = Nothing
    Exit Sub
ReporteAdeudadosVinculadosErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub


End Sub
'TORE - Automatizacion de Prorrogas
'Modificacion: Se agrego la columna Prorroga
Public Sub ReporteArendirCuentaViaticosLibro(pdFecha As Date, pdFecha2 As Date, psOpeCod As String, Optional pbCajaChica As Boolean = False)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsFechaDesde  As String
Dim lsFechaHasta As String
Dim Lineas As Long
Dim TotalDebe As Currency
Dim TotalHaber As Currency
Dim TotalImporte As Currency
Dim TotalSust As Currency
Dim TotalRend As Currency

Dim Total As Long
Dim j As Long
Dim lsObjARendir As String
Dim i As Integer
Dim lsFecha As String
Dim lnImporte As Currency
Dim lsHoja As String
Dim lsCodCtaCont As String
Dim lsImpre As String
Dim lnPaginas As Integer
Dim lsMsgErr As String
Dim lsFecRend As String
Dim nLin As Integer
Dim lcCta2 As String
Dim lcCodCta1 As String
Dim lcCodCta2 As String

On Error GoTo ErrReporteArendir
Set rs = CargaOpeCta(psOpeCod, "D")
If rs.EOF Then
    RSClose rs
    MsgBox "No se definió Cuenta Contable para Analizar Pendiente", vbInformation, "¡Aviso!"
    Exit Sub
End If
lsCodCtaCont = RSMuestraLista(rs)
'lsCodCtaCont = rs!cCtaContCod

If pbCajaChica Then
    lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoCajaChica & "," & gArendirTipoCajaGeneral & "," & gArendirTipoViaticos
Else
    'lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoCajaGeneral & "," & gArendirTipoViaticos
    lsObjARendir = gArendirTipoAgencias & "," & gArendirTipoViaticos
End If

lsArchivo = App.path & "\SPOOLER\LibAuxViaticos_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    'oBarra.CloseForm Me
    Exit Sub
End If

lsHoja = "LibAuxViaticos"

ExcelAddHoja lsHoja, xlLibro, xlHoja1

'Modificado PASI20140512 TI-ERS060-2014
'xlHoja1.Range(xlHoja1.Cells(1, 8), xlHoja1.Cells(1500, 8)).NumberFormat = "#,##0.00"
'xlHoja1.Range(xlHoja1.Cells(1, 10), xlHoja1.Cells(1500, 10)).NumberFormat = "#,##0.00"
'xlHoja1.Range(xlHoja1.Cells(1, 11), xlHoja1.Cells(1500, 11)).NumberFormat = "#,##0.00"
'xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1500, 1)).NumberFormat = "dd/mm/yyyy"
'xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(1500, 6)).NumberFormat = "dd/mm/yyyy"
'xlHoja1.Range(xlHoja1.Cells(1, 9), xlHoja1.Cells(1500, 9)).NumberFormat = "dd/mm/yyyy"

xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1500, 1)).NumberFormat = "dd/mm/yyyy"
xlHoja1.Range(xlHoja1.Cells(1, 5), xlHoja1.Cells(1500, 5)).NumberFormat = "dd/mm/yyyy"
xlHoja1.Range(xlHoja1.Cells(1, 7), xlHoja1.Cells(1500, 7)).NumberFormat = "dd/mm/yyyy"
xlHoja1.Range(xlHoja1.Cells(1, 8), xlHoja1.Cells(1500, 8)).NumberFormat = "dd/mm/yyyy"
xlHoja1.Range(xlHoja1.Cells(1, 9), xlHoja1.Cells(2500, 9)).NumberFormat = "dd/mm/yyyy" 'TORE - columnas Prorrogas
xlHoja1.Range(xlHoja1.Cells(1, 10), xlHoja1.Cells(1500, 10)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 11), xlHoja1.Cells(1500, 12)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 13), xlHoja1.Cells(1500, 13)).NumberFormat = "#,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, 14), xlHoja1.Cells(1500, 13)).NumberFormat = "#,##0.00"
'end PASI

lsImpre = ""
lnPaginas = 0
Linea lsImpre, PrnSet("MI", 5) & ReporteArendirCuentaViaticosLibroEncabezado(pdFecha, pdFecha2, Val(Mid(psOpeCod, 3, 1)), lnPaginas), 0

Lineas = 6
lsFechaDesde = Format(pdFecha, "yyyymmdd")
lsFechaHasta = Format(pdFecha2, "yyyymmdd")

'*** PEAC 20101111
'sql = "Select A.CMOVNRO, A.CMOVDESC, A.IMPORTE, A.nDocTpo, A.cDocNro, A.cPersNombre,A.IMPORTEME," _
'    & "      A.MovRend, A.ImporteSust, A.ImporteMESust, A.ImporteRend, A.ImporteMERend, Sum(MV.nMovViaticosDias) As nDias " _
'    & "From(SELECT   M.CMOVNRO, MR.NMOVNROREF, M.CMOVDESC, MC.NMOVIMPORTE IMPORTE, md.nDocTpo, md.cDocNro, " _
'    & "          P.cPersNombre, ISNULL(ME.NMOVMEIMPORTE,0) AS IMPORTEME, " _
'    & "          ISNULL(Max(ISNULL(Rend.cMovNro, Sust.cMovNro)),'') MovRend, ISNULL(SUM(Sust.nMovImporte),0) ImporteSust, ISNULL(SUM(Sust.nMovMEImporte),0) ImporteMESust, " _
'    & "          ISNULL(Rend.nMovImporte,0) ImporteRend, ISNULL(Rend.nMovMEImporte,0) ImporteMERend " _
'    & "     FROM     MOV M LEFT JOIN MOVDOC MD ON MD.nMovNro=M.nMovNro and NOT md.nDocTpo = " & TpoDocVoucherEgreso _
'    & "         INNER JOIN MOVCta MC ON M.nMovNro=MC.nMovNro " _
'    & "         LEFT JOIN MOVME ME  ON ME.nMovNro=MC.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "         INNER JOIN MOVRef MR ON MR.nMovNro=M.nMovNro " _
'    & "         INNER JOIN MOVArendir MO ON MO.nMovNro=MR.nMovNroRef " _
'    & "         INNER JOIN Persona P ON MO.cPersCOD= P.cPersCOD " _
'    & "         LEFT JOIN (SELECT m.cMovNro, mr.nMovNroRef, mc.nMovImporte* -1 nMovImporte, ISNULL(me.nMovMEImporte,0)*-1 nMovMEImporte " _
'    & "                     FROM Mov m JOIN MovCta mc ON m.nMovNro = mc.nMovNro " _
'    & "                         LEFT JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "                         JOIN MovRef mr on mr.nMovNro = m.nMovNro " _
'    & "                     WHERE m.nMovEstado in( " & gMovEstContabMovContable & ") and m.nMovFlag = " & gMovFlagVigente & " and mc.cCtaContCod IN (" & lsCodCtaCont & ", '29" & Mid(psOpeCod, 3, 1) & "80706' ) " _
'    & "                         and not m.cOpeCod LIKE '40__[356]%' " _
'    & "         ) Sust ON (Sust.nMovNroRef = m.nMovNro and mo.cTpoArendir = 1) or  (Sust.nMovNroRef = mo.nMovNro and mo.cTpoArendir = 2) "
'
'sql = sql & "    LEFT JOIN (SELECT m.cMovNro, mr.nMovNroRef, mc.nMovImporte* -1 nMovImporte, ISNULL(me.nMovMEImporte,0)*-1 nMovMEImporte " _
'    & "          FROM Mov m JOIN MovCta mc ON m.nMovNro = mc.nMovNro " _
'    & "            LEFT JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "            JOIN MovRef mr on mr.nMovNro = m.nMovNro " _
'    & "          WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente & " and  " _
'    & "            ( (mc.cCtaContCod in (" & lsCodCtaCont & ") and m.cOpeCod LIKE '40__5%') or (mc.cCtaContCod = '29" & Mid(psOpeCod, 3, 1) & "80706' and m.cOpeCod LIKE '40__6%') ) " _
'    & "    ) Rend ON (Rend.nMovNroRef = m.nMovNro and mo.cTpoArendir = 1) or  (Rend.nMovNroRef = mo.nMovNro and mo.cTpoArendir = 2) " _
'    & "WHERE MC.CCTACONTCOD in (" & lsCodCtaCont & ") AND MC.NMOVIMPORTE > 0 AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & lsFechaDesde & "' AND '" & lsFechaHasta & "' " _
'    & "  AND M.nMovFlag NOT IN(" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "  AND M.cOpeCod LIKE '40__3%' " _
'    & "  AND MO.cTpoArendir in (2) " _
'    & "GROUP BY M.CMOVNRO, MR.NMOVNROREF, M.CMOVDESC, MC.NMOVIMPORTE, md.nDocTpo, md.cDocNro, " _
'    & "         P.cPersNombre, ISNULL(ME.NMOVMEIMPORTE,0), " _
'    & "         ISNULL(Rend.nMovImporte,0), ISNULL(Rend.nMovMEImporte,0) ) AS A " _
'    & "Inner Join MovViaticos MV On A.nMovNroRef = MV.nViaticoMovNro  " _
'    & "Group By A.CMOVNRO, A.CMOVDESC, A.IMPORTE, A.nDocTpo, A.cDocNro, A.cPersNombre,A.IMPORTEME, A.MovRend, " _
'    & "    A.ImporteSust, A.ImporteMESust, A.ImporteRend, A.ImporteMERend " _
'    & "ORDER BY A.cMovNro"
    
    lcCta2 = "29" & Mid(psOpeCod, 3, 1) & "80706"
    lsCodCtaCont = Replace(lsCodCtaCont, "'", "")
    lcCodCta1 = lsCodCtaCont & "," & lcCta2
    lcCodCta2 = Replace(lcCodCta1, "'", "")
    sql = " EXEC stp_sel_ReporteArendirCuentaViaticosLibro '" & lsCodCtaCont & "','" & lcCta2 & "','" & lcCodCta2 & "','" & lsFechaDesde & "','" & lsFechaHasta & "'"
    
'*** FIN PEAC
       
    
TotalImporte = 0
TotalSust = 0
TotalRend = 0
Set oBarra = New clsProgressBar
ProgressShow oBarra, frmReportes, ePCap_CaptionPercent, True
oBarra.Progress 0, "A Rendir Cuenta : Libro Auxiliar", "Cargando datos...", "", vbBlue

Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
Total = rs.RecordCount
nLin = 7
oBarra.Progress 1, "A Rendir Cuenta Viaticos: Libro Auxiliar", "Cargando datos...", "", vbBlue
oBarra.Max = Total
Do While Not rs.EOF
    
    oBarra.Progress rs.Bookmark, "A Rendir Cuenta Viaticos: Libro Auxiliar", "Generando Reporte...", "", vbBlue
    j = j + 1
    lnImporte = IIf(Mid(psOpeCod, 3, 1) = "1", rs!Importe, rs!ImporteME)
    lsFecha = Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4)
    If rs!MovRend = "" Then
        lsFecRend = ImpreFormat("", 10)
    Else
        lsFecRend = ImpreFormat(GetFechaMov(rs!MovRend, True), 10)
    End If
    Linea lsImpre, ImpreFormat(Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), 12, 0) & _
            ImpreFormat(Format(rs!nDocTpo, "00") & " " & rs!cDocNro, 14) & ImpreFormat(PstaNombre(rs!cPersNombre, False), 32) & ImpreFormat(rs!cMovDesc, 40) & _
            ImpreFormat(rs!nDias, 10) & ImpreFormat(rs!plazo_rendir, 12, 0) & ImpreFormat(rs!Plazo_Rendir_Prorroga, 15, 0) & ImpreFormat(IIf(rs!Dias_atraso < 0, 0#, rs!Dias_atraso), 10) & IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, ImpreFormat(rs!ImporteME, 10) & _
               lsFecRend & ImpreFormat(rs!ImporteMESust, 10) & ImpreFormat(rs!ImporteMERend, 10), _
               ImpreFormat(rs!Importe, 10) & ImpreFormat(lsFecRend, 10) & _
               ImpreFormat(rs!ImporteSust, 10) & ImpreFormat(rs!ImporteRend, 10))
    
    'Modificado PASI20140512 TI-ERS060-2014
'    xlHoja1.Cells(nLin, 1) = "'" & lsFecha
'    xlHoja1.Cells(nLin, 2) = Format(rs!nDocTpo, "00") & " " & rs!cDocNro
'    xlHoja1.Cells(nLin, 3) = PstaNombre(rs!cPersNombre, False)
'    xlHoja1.Cells(nLin, 4) = ImpreFormat(rs!cMovDesc, 180)
'    xlHoja1.Cells(nLin, 5) = rs!ndias
'    '*** PEAC 20101111
'    xlHoja1.Cells(nLin, 6) = "'" & Format(rs!plazo_rendir, "dd/mm/yyyy")
'    xlHoja1.Cells(nLin, 7) = IIf(rs!Dias_atraso < 0, 0#, rs!Dias_atraso)
'    '*** FIN PEAC
'    xlHoja1.Cells(nLin, 8) = lnImporte
'    xlHoja1.Cells(nLin, 9) = lsFecRend
'    xlHoja1.Cells(nLin, 10) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteSust, rs!ImporteMESust)
'    xlHoja1.Cells(nLin, 11) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteRend, rs!ImporteMERend)
    
    xlHoja1.Cells(nLin, 1) = "'" & lsFecha
    xlHoja1.Cells(nLin, 2) = Format(rs!nDocTpo, "00") & " " & rs!cDocNro
    xlHoja1.Cells(nLin, 3) = PstaNombre(rs!cPersNombre, False)
    xlHoja1.Cells(nLin, 4) = ImpreFormat(rs!cMovDesc, 180)
    xlHoja1.Cells(nLin, 5) = rs!dFecPartida
    xlHoja1.Cells(nLin, 6) = rs!nDias
    xlHoja1.Cells(nLin, 7) = rs!dFecLlegada
    xlHoja1.Cells(nLin, 8) = "'" & Format(rs!plazo_rendir, "dd/mm/yyyy")
    xlHoja1.Cells(nLin, 9) = "'" & Format(rs!Plazo_Rendir_Prorroga, "dd/mm/yyyy") 'TORE - Automatizacion de las ampliaciones de prorrogas.
    xlHoja1.Cells(nLin, 10) = IIf(rs!Dias_atraso < 0, 0#, rs!Dias_atraso)
    xlHoja1.Cells(nLin, 11) = lnImporte
    xlHoja1.Cells(nLin, 12) = lsFecRend
    xlHoja1.Cells(nLin, 13) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteSust, rs!ImporteMESust)
    xlHoja1.Cells(nLin, 14) = IIf(Mid(psOpeCod, 3, 1) = "1", rs!ImporteRend, rs!ImporteMERend)
    'end PASI
    
    Lineas = Lineas + 1
    If Lineas > 60 Then
        Linea lsImpre, ReporteArendirCuentaViaticosLibroEncabezado(pdFecha, pdFecha2, Val(Mid(psOpeCod, 3, 1)), lnPaginas), 0 'PASI20140611
        Lineas = 6
    End If
    TotalImporte = TotalImporte + rs!Importe
    TotalSust = TotalSust + rs!ImporteSust
    TotalRend = TotalRend + rs!ImporteRend
    nLin = nLin + 1
    rs.MoveNext
    DoEvents
Loop

'Modificado PASI20140512 TI-ERS060-2014
'xlHoja1.Cells(nLin + 1, 7) = "TOTALES :"
'xlHoja1.Cells(nLin + 1, 8) = TotalImporte
'xlHoja1.Cells(nLin + 1, 10) = TotalSust
'xlHoja1.Cells(nLin + 1, 11) = TotalRend

xlHoja1.Cells(nLin + 1, 10) = "TOTALES :"
xlHoja1.Cells(nLin + 1, 11) = TotalImporte
xlHoja1.Cells(nLin + 1, 13) = TotalSust
xlHoja1.Cells(nLin + 1, 14) = TotalRend

RSClose rs
ProgressClose oBarra, frmReportes, True
Set oBarra = Nothing
Linea lsImpre, String(220, "=")
Linea lsImpre, ImpreFormat("TOTALES :", 10, 102) & ImpreFormat(TotalImporte, 15, 2) & space(11) & ImpreFormat(TotalSust, 11, 2) & ImpreFormat(TotalRend, 10, 2)
If lsImpre <> "" Then
    EnviaPrevio lsImpre, "A RENDIR CUENTA VIATICOS: LIBRO AUXILIAR", gnLinPage, True
End If

ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    
Exit Sub
ErrReporteArendir:
lsMsgErr = Err.Description
    ProgressClose oBarra, frmReportes, True
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'*** PEAC 20100510
Public Sub ReporteChequesPorCobrarConvenio(pdFecha As Date, psOpeCod As String)
    Dim lFecha As String
    Dim lnTotal As Currency
    Dim sTexto  As String
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim oCon As DConecta
    Dim oOpe As DOperacion
    Dim lnLin As Integer
    Dim lnPag As Integer
    Dim lnTotal1 As Currency, lnTotal2 As Currency, lnTotal3 As Currency
    Set oCon = New DConecta
    oCon.AbreConexion

    lFecha = pdFecha
    lnTotal = 0
    lnPag = 0
    lnLin = gnLinPage
    sTexto = ""
    
    sSql = " exec stp_sel_ObtieneChequesPorCobrarConvenio '" & Format(pdFecha, "yyyymmdd") & "','" & Mid(psOpeCod, 3, 1) & "' "

    Set rs = oCon.CargaRecordSet(sSql)
    
    If (rs.EOF And rs.BOF) Then
        rs.Close
        MsgBox "No se encontraron datos para este Reporte", vbInformation, "Error"
        Exit Sub
    End If
    
    lnTotal1 = 0:   lnTotal2 = 0:    lnTotal3 = 0
    sTexto = PrnSet("C+") + PrnSet("B+")
    Do While Not rs.EOF
       If lnLin > gnLinPage - 4 Then
          Linea sTexto, CabeRepo(gsNomCmac, gsNomAge, "Finanzas", "", Format(gdFecSis, gsFormatoFechaView), "REPORTE DE CHEQUES POR COBRAR CONVENIO " & " EN " & IIf(Mid(psOpeCod, 3, 1) = "1", "M.N.", "M.E."), " A la Fecha : " & lFecha, "", "", lnPag, gnColPage), 0, lnLin
          lnLin = 6
          Linea sTexto, String(145, "="), , lnLin
          Linea sTexto, ImpreFormat("AGENCIA", 25) & ImpreFormat("ENTIDAD", 30) & ImpreFormat("FECHA EMISION", 15) & ImpreFormat("NUM. CHEQUE", 15) & ImpreFormat("IMPORTE INICIAL", 15) & ImpreFormat("MONTO USADO", 15) & ImpreFormat("SALDO", 15), , lnLin
          Linea sTexto, String(145, "-") + PrnSet("B-"), , lnLin
       End If
       lnTotal1 = lnTotal1 + rs!Importe_Inicial
       lnTotal2 = lnTotal2 + rs!Monto_Usado
       lnTotal3 = lnTotal3 + rs!Importe_Inicial - rs!Monto_Usado
       Linea sTexto, ImpreFormat(rs!Agencia, 25) & ImpreFormat(rs!Entidad, 30) & ImpreFormat(rs!Fecha, 12) & ImpreFormat(rs!Num_Cheque, 18, , True) & ImpreFormat(rs!Importe_Inicial, 12, , True) & ImpreFormat(rs!Monto_Usado, 12, , True) & ImpreFormat(rs!Importe_Inicial - rs!Monto_Usado, 12, , True), , lnLin
    rs.MoveNext
    Loop
    RSClose rs
    sTexto = sTexto & oImpresora.gPrnSaltoLinea
    Linea sTexto, String(145, "-") + PrnSet("B-"), , lnLin
    sTexto = sTexto + PrnSet("B+") + ImpreFormat("", 83) & ImpreFormat(lnTotal1, 12, 2, True) & ImpreFormat(lnTotal2, 12, 2, True) & ImpreFormat(lnTotal3, 12, 2, True) & oImpresora.gPrnSaltoLinea & PrnSet("C-")
    EnviaPrevio sTexto, "Reporte de Cheques Por Cobrar Convenio", gnLinPage, False
End Sub
