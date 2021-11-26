Attribute VB_Name = "gCaptaReport"
Option Explicit

Dim lnItem As Long
Dim lnPagina As Long
Const lnRango = 59

'Private Sub QuitarSalto(rtf As String)
'    If Right(rtf, 1) = oImpresora.gPrnSaltoPagina  Then
'        rtf = Mid(rtf, 1, Len(rtf) - 1)
'    End If
'End Sub
'
'Private Sub PonerSalto(rtf As String)
'    If rtf = "" Then Exit Sub
'    If Right(rtf, 1) <> oImpresora.gPrnSaltoPagina  Then
'        rtf = rtf & oImpresora.gPrnSaltoPagina
'    End If
'End Sub
'
'Public Function GetNomOpe(psCodOpe As String, Optional pNum As Integer = 22) As String
'    Dim sqlOpe As String
'    Dim rsOpe As New ADODB.Recordset
'    Dim lsNomOpe As String
'
'    sqlOpe = "Select cNomOpe from " & gcCentralCom & "Operacion where cCodOpe = '" & psCodOpe & "'"
'    'rsOpe.Open sqlOpe, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If rsOpe.BOF And rsOpe.EOF Then
'        GetNomOpe = " "
'    Else
'        lsNomOpe = Left(rsOpe!cNomOpe, pNum)
'        GetNomOpe = lsNomOpe
'        rsOpe.Close
'    End If
'
'    Set rsOpe = Nothing
'End Function
'
'
''************************************************************************
''PLAZO FIJO
''********************************************************************************************
'Public Sub ReporteMovPF(pRich As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'
'    Dim SQL1 As String
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lnDep As Double
'    Dim lnRet As Double
'
'    Dim lnTotDep As Double
'    Dim lnTotRet As Double
'    Dim lsCebecera As String
'
'    Dim lsTEM As String
'    Dim I As Long
'    Dim lsAdd As String
'    lnTotRet = 0
'    lnTotDep = 0
'    lsAdd = IIf(psFecha = "", "", " Del " & Format$(psFecha, gsformatofechaview))
'    If psCodUser = "XXX" Then
'        'pRich = pRich + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsEmpresa, gdFecSis, psMoneda)
'    Else
'        'pRich = pRich + CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, psMoneda)
'        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, gsNomAge, gsEmpresa, gdFecSis, psMoneda)
'    End If
'
'    pnPagina = pnPagina - 1
'
'    lsTEM = pRich
'
'    pRich = ""
'
'    Call ReporteAperturasPF(pRich, "Apertura de Cuentas PlazoFijo-Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoEfectivoPF(pRich, "Movimientos en Efectivo de Cuentas de Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call ReporteAperturasPFChq(pRich, "Apertura de Cuentas Plazo Fijo-Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoCancelacionPF(pRich, "Cancelación de Cuentas Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'
'    If lnTotDep <> 0 Or lnTotRet <> 0 Then
'        pRich = lsCebecera + pRich
'        QuitarSalto pRich
'        pRich = pRich + Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
'        pRich = pRich + Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
'        pRich = pRich & oImpresora.gPrnSaltoPagina
'    End If
'
'    lnTotRet = 0
'    lnTotDep = 0
'
'    Call MovimientoExternosPF(pRich, "Cuentas de -Plazo Fijo- Movidas por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call CancelacionExternoPF(pRich, "Cuentas Canceladas de -Plazo Fijo- por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoTramitePF(pRich, "Cuentas de -Plazo Fijo- Movidas en otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha)
'
'    pRich = lsTEM + pRich
'    lsTEM = ""
'End Sub
'
''Movimiento que se realiza con cuentas de otra agencia desde
''este otro servidor
'Public Function MovimientoTramitePF(pRich As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'    Dim lbBan As Boolean
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepAge As Double
'    Dim lnTRetAge As Double
'
'    Dim I As Long
'
'    Dim lsGru As String
'
'    lnTRet = 0
'    lnTDep = 0
'    lnTRetAge = 0
'    lnTDepAge = 0
'    lbBan = False
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            'SQL1 = "Select cCodAge, dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from Trandiaria A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsPFOTRetInt & "','" & gsPFOTCanAct & "') and cFlag is null order by A.cCodCta"
'        Else
'            'SQL1 = "Select cCodAge, dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from TrandiariaConsol A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsPFOTRetInt & "','" & gsPFOTCanAct & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        'SQL1 = "Select cCodAge, dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from Trandiaria A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsPFOTRetInt & "','" & gsPFOTCanAct & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta "
'    End If
'
'    'RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If Not RSVacio(RegTrandiaria) Then
'
'        pRich = pRich & CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'        pRich = pRich & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Operacion;30;Deposito;20;Retiro;20;Usu.R;6;Hora;12;", pnItem)
'
'        lsGru = ""
'
'        While Not RegTrandiaria.EOF
'            lbBan = True
'
'            If lsGru <> "" And ((lsGru <> Mid(RegTrandiaria!cCodCta, 1, 2) And Len(RegTrandiaria!cCodCta) = 12) Or (lsGru <> "0" & Mid(RegTrandiaria!cCodCta, 3, 1) And Len(RegTrandiaria!cCodCta) = 8)) Then
'                pRich = pRich + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";67;" & Format(lnTRetAge, "###,##0.00") & ";20; ;18;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'
'            If Len(RegTrandiaria!cCodCta) = 12 Then
'                If lsGru <> Mid(RegTrandiaria!cCodCta, 1, 2) Then
'                    lsGru = Mid(RegTrandiaria!cCodCta, 1, 2)
'                    pRich = pRich + Encabezado("Agencia: " & GetNomAge("112" & lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
'                If lsGru <> "0" & Mid(RegTrandiaria!cCodCta, 3, 1) Then
'                    lsGru = "0" & Mid(RegTrandiaria!cCodCta, 3, 1)
'                    pRich = pRich + Encabezado("Agencia: " & GetNomAge("112" & lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnTRetAge = lnTRetAge + (RegTrandiaria!nMonTran * -1)
'                If RegTrandiaria!cCodOpe = gsPFOTCanAct Then
'                    lnTRet = lnTRet + GetMonIntCanOT(RegTrandiaria!cCodCta)
'                    lnTRetAge = lnTRetAge + GetMonIntCanOT(RegTrandiaria!cCodCta)
'                End If
'            End If
'
'            pRich = pRich & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(5) & Trim(GetNomOpe(RegTrandiaria!cCodOpe)) _
'                               & Space(25 - Len(Trim(GetNomOpe(RegTrandiaria!cCodOpe)))) & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(6 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                               & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & oImpresora.gPrnSaltoLinea
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRich = pRich & oImpresora.gPrnSaltoPagina
'                pRich = pRich + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'                pRich = pRich + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Operacion;30;Deposito;20;Retiro;20;Usu.R;6;Hora;12;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        If lbBan Then
'            pRich = pRich + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";67;" & Format(lnTRetAge, "###,##0.00") & ";20; ;18;", pnItem)
'        End If
'
'        pRich = pRich + Encabezado("Total Resumen;15;" & Format(lnTDep, "###,##0.00") & ";67;" & Format(lnTRet, "###,##0.00") & ";20; ;18;", pnItem)
'
'        pRich = pRich & oImpresora.gPrnSaltoPagina
'        pnItem = pnItem + 1
'
'        RegTrandiaria.Close
'    End If
'
'    Set RegTrandiaria = Nothing
'
'End Function
'
'
'Public Function CancelacionExternoPF(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'
'    Dim lsGru As String
'    Dim lnTRetAge As Double
'    Dim lnTDepAge As Double
'    Dim lbBan As Boolean
'
'    Dim I As Long
'
'    lnTRetAge = 0
'    lnTDepAge = 0
'
'    lnTRet = 0
'    lnTDep = 0
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select cCodAge, dFecTran, nPlazo, A.cCodUsu, cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOACanAct & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select cCodAge, dFecTran, nPlazo, A.cCodUsu, cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nTasaIntPF, nSaldCnt from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOACanAct & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'
'    Else
'        SQL1 = "Select cCodAge, dFecTran, nPlazo, A.cCodUsu, cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOACanAct & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'
'    Else
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;20;Can.Int.;20;Can.Cta;20;Usu.;9;Hora;15;TInt;10;D.Plazo;10;", pnItem)
'
'        While Not RegTrandiaria.EOF
'
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                pRTF = pRTF + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";39;" & Format(lnTRetAge, "###,##0.00") & ";20; ;44;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If Len(RegTrandiaria!cCodCta) = 12 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'
'            lnTDepAge = lnTDepAge + GetMonIntCanOA(RegTrandiaria!cCodCta)
'            lnTRetAge = lnTRetAge + (RegTrandiaria!nMonTran * -1)
'
'
'            lsDep = Format(GetMonIntCanOA(RegTrandiaria!cCodCta), "###,##0.00")
'            lnTDep = lnTDep + GetMonIntCanOA(RegTrandiaria!cCodCta)
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(9 - Len(RegTrandiaria!cCodusurem)) & RegTrandiaria!cCodusurem _
'                               & Space(15 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(10 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                               & Space(10 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'            'Str(RegTrandiaria!nPlazo)
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Can.Int.;20;Can.Cta;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;10;D.Plazo;10;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        lnTDepT = lnTDep
'        lnTRetT = lnTRet
'
'        If lbBan Then pRTF = pRTF + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";39;" & Format(lnTRetAge, "###,##0.00") & ";20; ;44;", pnItem)
'
'
'        RegTrandiaria.Close
'
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTDepT <> 0 Or lnTRetT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";39;" & Trim(Format(lnTRetT, "###,##0.00")) & ";20; ;44;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRetT + lnTDepT
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
''Operacioens hechas por otras agencias en este servidor
'Public Function MovimientoExternosPF(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'    Dim lnValor As Double
'    Dim lnOpe As Double
'
'    Dim lsGru As String
'
'
'    Dim lbBan As Boolean
'
'    Dim lnTRetAge As Double
'    Dim lnTDepAge As Double
'    Dim lsCodUsuX As String
'
'
'    lnTRetAge = 0
'    lnTDepAge = 0
'
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'
'    'and A.cCodUsu = '" & psCodUser & "'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, cCodAge, nPlazo, A.cCodUsu, cCodUsuRem , cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOARetInt & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, cCodAge, nPlazo, A.cCodUsu, cCodUsuRem , cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOARetInt & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, cCodAge, nPlazo, A.cCodUsu, cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFOARetInt & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.BOF And RegTrandiaria.EOF Then
'
'    Else
'
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;IntDevAnt;20;Retiro;20;Usu.;10;Hora;17;TInt;9;Plazo;10;", pnItem)
'
'        lsGru = ""
'        lbBan = False
'
'        While Not RegTrandiaria.EOF
'
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                pRTF = pRTF + Encabezado("Total Agencia;15; ;37;" & Format(lnTRetAge * -1, "###,##0.00") & ";20; ;46;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If Len(RegTrandiaria!cCodCta) = 12 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = Trim(RegTrandiaria!cNumDoc)
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = RegTrandiaria!nMonTran
'                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnOpe = RegTrandiaria!nMonTran
'                lnTRetAge = lnTRetAge + RegTrandiaria!nMonTran
'            End If
'
'            If IsNull(RegTrandiaria!nIntDev) Then
'                lnValor = 0
'            Else
'                lnValor = RegTrandiaria!nIntDev
'            End If
'
'            If IsNull(RegTrandiaria!cCodUsu) Then
'                If IsNull(RegTrandiaria!cCodusurem) Then
'                    lsCodUsuX = ""
'                Else
'                    lsCodUsuX = RegTrandiaria!cCodusurem
'                End If
'            Else
'                lsCodUsuX = RegTrandiaria!cCodUsu
'            End If
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(10 - Len(lsCodUsuX)) & lsCodUsuX _
'                               & Space(17 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                               & Space(11 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'            'Str(RegTrandiaria!nPlazo)
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;IntDevAnt;20;Retiro;20;Usu.;10;Hora;17;TInt;9;Plazo;10;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        If lbBan Then pRTF = pRTF + Encabezado("Total Agencia;15; ;37;" & Format(lnTRetAge * -1, "###,##0.00") & ";20; ;46;", pnItem)
'
'        lnTRetT = lnTRet
'        lnTDepT = lnTDep
'
'        RegTrandiaria.Close
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTRetT <> 0 Or lnTDepT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15; ;37;" & Format(lnTRetT, "###,##0.00") & ";20; ;46;", pnItem)
'        pnTotRet = pnTotRet + lnTRetT
'        pnTotDep = pnTotDep + lnTDepT
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'Public Function ReporteAperturasPF(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim I As Long
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnValor As Double
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeEfe & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeEfe & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeEfe & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'        Exit Function
'    End If
'
'    pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'    pRTF = pRTF + Encabezado("Nro Cuenta;15;Nro.Doc.;20;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;9;D.Plazo;10;", pnItem)
'
'    While Not RegTrandiaria.EOF
'
'        If IsNull(RegTrandiaria!cNumDoc) Then
'            lsDoc = " "
'        Else
'            lsDoc = RegTrandiaria!cNumDoc
'        End If
'
'        If RegTrandiaria!nMonTran >= 0 Then
'            lnTDep = lnTDep + RegTrandiaria!nMonTran
'        Else
'            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'        End If
'
'        If IsNull(RegTrandiaria!nSaldCnt) Then
'            lnValor = 0
'        Else
'            lnValor = RegTrandiaria!nSaldCnt
'        End If
'
'        pRTF = pRTF & Space(15 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                           & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                           & Space(20 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
'                           & Space(9 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                           & Space(15 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                           & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
'                           & Space(9 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                           & Space(10 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'        pnItem = pnItem + 1
'
'        If pnItem = lnRango Then
'            pRTF = pRTF & oImpresora.gPrnSaltoPagina
'            pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'            pRTF = pRTF + Encabezado("Nro Cuenta;15;Nro.Doc.;20;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;9;D.Plazo;10", pnItem)
'        End If
'
'        RegTrandiaria.MoveNext
'    Wend
'
'    If lnTRet <> 0 Or lnTDep <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDep, "###,##0.00")) & ";40; ;63;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRet
'        pnTotDep = pnTotDep + lnTDep
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'End Function
'
'Public Function MovimientoEfectivoPF(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'    Dim lnValor As Double
'    Dim lnOpe As Double
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsu, nPlazo, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFRetIntNC & "','" & gsPFRetInt & "') and cFlag is null  order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsu, nPlazo, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFRetIntNC & "','" & gsPFRetInt & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsu, nPlazo, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev, nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFRetIntNC & "','" & gsPFRetInt & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.BOF And RegTrandiaria.EOF Then
'
'    Else
'
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;20;SalAnt;20;Retiro;20;Usu.;8;Hora;15;TInt;10;D.Plazo;11;", pnItem)
'
'        While Not RegTrandiaria.EOF
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnOpe = RegTrandiaria!nMonTran
'            End If
'
'            If IsNull(RegTrandiaria!nSaldCnt) Then
'                lnValor = 0
'            Else
'                lnValor = RegTrandiaria!nSaldCnt
'            End If
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(8 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                               & Space(15 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(10 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                               & Space(11 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;20;SalAnt;20;Retiro;20;Usu.;8;Hora;15;TInt;10;D.Plazo;11;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        pRTF = pRTF + Encabezado("Resumen;15;" & Format(lnTRet, "###,##0.00") & ";59; ;44;", pnItem)
'
'        lnTRetT = lnTRet
'        lnTDepT = lnTDep
'
'        RegTrandiaria.Close
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTRetT <> 0 Or lnTDepT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Format(lnTRetT, "###,##0.00") & ";59; ;44;", pnItem)
'        pnTotRet = pnTotRet + lnTRetT
'        pnTotDep = pnTotDep + lnTDepT
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'Public Function ReporteAperturasPFChq(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim I As Long
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnValor As Double
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            If psFecha = "" Then
'                SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeChq & "') and cFlag is null order by A.cCodCta"
'            Else
'                SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeChq & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'            End If
'        Else
'            SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeChq & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntPF from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFApeChq & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'        Exit Function
'    End If
'
'    pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'    pRTF = pRTF + Encabezado("Nro Cuenta;15;Nro.Doc.;18;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;11;D.Plazo;10;", pnItem)
'
'    While Not RegTrandiaria.EOF
'
'        If IsNull(RegTrandiaria!cNumDoc) Then
'            lsDoc = " "
'        Else
'            lsDoc = RegTrandiaria!cNumDoc
'        End If
'
'        If RegTrandiaria!nMonTran >= 0 Then
'            lnTDep = lnTDep + RegTrandiaria!nMonTran
'        Else
'            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'        End If
'
'        If IsNull(RegTrandiaria!nSaldCnt) Then
'            lnValor = 0
'        Else
'            lnValor = RegTrandiaria!nSaldCnt
'        End If
'
'        pRTF = pRTF & Space(15 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                           & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                           & Space(20 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
'                           & Space(9 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                           & Space(15 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                           & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
'                           & Space(11 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                           & Space(10 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'        pnItem = pnItem + 1
'
'        If pnItem = lnRango Then
'            pRTF = pRTF & oImpresora.gPrnSaltoPagina
'            pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'            pRTF = pRTF + Encabezado("Nro Cuenta;15;Nro.Doc.;18;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;11;D.Plazo;10;", pnItem)
'        End If
'
'        RegTrandiaria.MoveNext
'    Wend
'
'    If lnTRet <> 0 Or lnTDep <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDep, "###,##0.00")) & ";38; ;65;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRet
'        pnTotDep = pnTotDep + lnTDep
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'End Function
'
'Public Function MovimientoCancelacionPF(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'    'and A.cCodUsu = '" & psCodUser & "'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev as nSaldDispCTS , nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFCanAct & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev as nSaldDispCTS , nTasaIntPF, nSaldCnt from TrandiariaConsol A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFCanAct & "') and cFlag is NULL And DateDiff(dd,dFectran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, nPlazo, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nIntDev as nSaldDispCTS , nTasaIntPF, nSaldCnt from Trandiaria A, PlazoFijo B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsPFCanAct & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'
'    Else
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;20;Can.Int.;20;Can.Cta;20;Usu.;9;Hora;15;TInt;10;D.Plazo;10;", pnItem)
'
'        While Not RegTrandiaria.EOF
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran <= 0 Then
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'            End If
'
'            lsDep = Format(GetMonIntCan(RegTrandiaria!cCodCta), "###,##0.00")
'            lnTDep = lnTDep + GetMonIntCan(RegTrandiaria!cCodCta)
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(9 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                               & Space(15 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(10 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
'                               & Space(10 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
'
'            'Str(RegTrandiaria!nPlazo)
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;20;Can.Int.;20;Can.Cta;20;Usu.;9;Hora;15;TInt;10;D.Plazo;10;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        lnTDepT = lnTDep
'        lnTRetT = lnTRet
'
'        RegTrandiaria.Close
'
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTDepT <> 0 Or lnTRetT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";39;" & Trim(Format(lnTRetT, "###,##0.00")) & ";20; ;44;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRetT + lnTDepT
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
''************************************************************************
''CTS
''********************************************************************************************
'Public Sub ReporteMovCTS(pRich As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'
'    Dim SQL1 As String
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lnDep As Double
'    Dim lnRet As Double
'
'    Dim lnTotDep As Double
'    Dim lnTotRet As Double
'    Dim lsCebecera As String
'
'    Dim lsTEM As String
'    Dim I As Long
'    Dim lsAdd As String
'
'    lsAdd = IIf(psFecha = "", "", " Del " & Format$(psFecha, gsformatofechaview))
'
'    lnTotRet = 0
'    lnTotDep = 0
'
'    If psCodUser = "XXX" Then
'        'pRich = pRich + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'    Else
'        'pRich = pRich + CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, psMoneda)
'        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, psMoneda)
'    End If
'
'    pnPagina = pnPagina - 1
'
'    lsTEM = pRich
'
'    pRich = ""
'
'    Call ReporteAperturasCTS(pRich, "Apertura de Cuentas CTS en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoEfectivoCTS(pRich, "Movimientos en Efectivo de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call ReporteAperturasCTSChq(pRich, "Apertura de Cuentas CTS con Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoCancelacionCTS(pRich, "Cancelación de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoChequeCTS(pRich, "Depositos con Cheque en Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'
'    If lnTotRet <> 0 Or lnTotDep <> 0 Then
'        pRich = lsCebecera + pRich
'        QuitarSalto pRich
'        pRich = pRich + Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
'        pRich = pRich + Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
'        pRich = pRich & oImpresora.gPrnSaltoPagina
'    End If
'
'    Call MovimientoExternosCTS(pRich, "Cuentas CTS Movidas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call CancelacionExternoCTS(pRich, "Cuentas CTS Canceladas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    Call MovimientoTramiteCTS(pRich, "Cuentas CTS Movidas en otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha)
'
'    pRich = lsTEM + pRich
'    lsTEM = ""
'End Sub
'
''Operacioens hechas por otras agencias en este servidor
'Public Function MovimientoExternosCTS(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'    Dim lnValor As Double
'    Dim lnOpe As Double
'    Dim lnPorRet As Double
'    Dim lbBan  As Boolean
'    Dim lsGru  As String
'    Dim lnTRetAge As Currency
'    Dim lnTDepAge As Currency
'
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt, ccodage from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.cCodCta = B.cCodCta and ccodope in ('" & gsCTSOARetEfe & "','" & gsCTSOADepEfe & "','" & gsCTSOADepChq & "') and cFlag is null order by ccodage, A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt, ccodage from TrandiariaConsol A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.cCodCta = B.cCodCta and ccodope in ('" & gsCTSOARetEfe & "','" & gsCTSOADepEfe & "','" & gsCTSOADepChq & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by ccodage, A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsuRem, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt, ccodage from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.cCodCta = B.cCodCta and ccodope in ('" & gsCTSOARetEfe & "','" & gsCTSOADepEfe & "','" & gsCTSOADepChq & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by ccodage, A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.BOF And RegTrandiaria.EOF Then
'
'    Else
'
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
'
'        lnPorRet = (ReadParametros("23110") / 100)
'
'        lsGru = ""
'        lnTRetAge = 0
'        lnTDepAge = 0
'
'        While Not RegTrandiaria.EOF
'
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                pRTF = pRTF + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;46;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If Len(RegTrandiaria!cCodCta) = 12 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                    lbBan = True
'                End If
'            End If
'
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = RegTrandiaria!nMonTran
''                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnOpe = RegTrandiaria!nMonTran
''                lnTRetAge = lnTRetAge + (RegTrandiaria!nMonTran * -1)
'            End If
'
'            If IsNull(RegTrandiaria!nSaldCnt) Then
'                lnValor = 0
'            Else
'                lnValor = RegTrandiaria!nSaldCnt
'            End If
'
'            If Trim(RegTrandiaria!cCodOpe) = gsCTSOARetInt Then
'                pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                                   & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                                   & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
'                                   & Space(20 - Len(lsDep)) & lsDep _
'                                   & Space(20 - Len(lsRet)) & lsRet _
'                                   & Space(6 - Len(RegTrandiaria!cCodusurem)) & RegTrandiaria!cCodusurem _
'                                   & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                                   & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'            Else
'                pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                                   & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                                   & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
'                                   & Space(20 - Len(lsDep)) & lsDep _
'                                   & Space(20 - Len(lsRet)) & lsRet _
'                                   & Space(6 - Len(RegTrandiaria!cCodusurem)) & RegTrandiaria!cCodusurem _
'                                   & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                                   & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'            End If
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;Sald.Cnt.;20;TInt;8;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        If lbBan Then pRTF = pRTF + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;46;", pnItem)
'        lnTRetT = lnTRet
'        lnTDepT = lnTDep
'
'        RegTrandiaria.Close
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTRetT <> 0 Or lnTDepT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Format(lnTDepT, "###,##0.00") & ";57;" & Format(lnTRetT, "###,##0.00") & ";20; ;25;", pnItem)
'        pnTotRet = pnTotRet + lnTRetT
'        pnTotDep = pnTotDep + lnTDepT
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'
''cancelacion externa CTS
'Public Function CancelacionExternoCTS(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'    Dim lbBan  As Boolean
'
'    Dim I As Long
'    Dim lsGru As String
'
'    Dim lnTDepAge As Currency
'    Dim lnTRetAge As Currency
'    lnTRet = 0
'    lnTDep = 0
'
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsu, A.cCodUsuRem, cCodAge, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSOACanAct & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsu, A.cCodUsuRem, cCodAge, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from TrandiariaConsol A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSOACanAct & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsu, A.cCodUsuRem, cCodAge, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSOACanAct & "') and cFlag is null and a.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'
'    Else
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Can.Int.;20;Can.Cta;20;Usu.;6;Hora;12;Sald.Disp.;20;TInt;8;", pnItem)
'
'        While Not RegTrandiaria.EOF
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                pRTF = pRTF + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";39;" & Format(lnTRetAge, "###,##0.00") & ";20; ;44;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If lsGru <> RegTrandiaria!cCodAge Then
'                lsGru = RegTrandiaria!cCodAge
'                pRTF = pRTF + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                lbBan = True
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'
'            lsDep = Format(GetMonIntCanOA(RegTrandiaria!cCodCta), "###,##0.00")
'            lnTDep = lnTDep + GetMonIntCanOA(RegTrandiaria!cCodCta)
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(6 - Len(RegTrandiaria!cCodusurem)) & RegTrandiaria!cCodusurem _
'                               & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(20 - Len("0.00")) & "0.00" _
'                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Can.Int.;20;Can.Cta;20;Usu.;6;Hora;12;Sald.Disp.;20;TInt;8;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        lnTDepT = lnTDep
'        lnTRetT = lnTRet
'
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";37;" & Trim(Format(lnTRetT, "###,##0.00")) & ";20; ;46;", pnItem)
'
'        RegTrandiaria.Close
'
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTDepT <> 0 Or lnTRetT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";37;" & Trim(Format(lnTRetT, "###,##0.00")) & ";20; ;46;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRetT + lnTDepT
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'
''Movimiento que se realiza con cuentas de otra agencia desde
''este otro servidor
'Public Function MovimientoTramiteCTS(pRich As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'    Dim lbBan As Boolean
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepAge As Double
'    Dim lnTRetAge As Double
'
'    Dim I As Long
'
'    Dim lsGru As String
'
'    lnTRet = 0
'    lnTDep = 0
'    lnTRetAge = 0
'    lnTDepAge = 0
'    lbBan = False
'
'    'and cCodUsu = '" & psCodUser & "'"
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, ccodage, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from Trandiaria A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsCTSOTRetEfe & "','" & gsCTSOTDepEfe & "','" & gsCTSOTDepChq & "','" & gsCTSOTRetInt & "','" & gsCTSOTCanAct & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, ccodage, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from TrandiariaConsol A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsCTSOTRetEfe & "','" & gsCTSOTDepEfe & "','" & gsCTSOTDepChq & "','" & gsCTSOTRetInt & "','" & gsCTSOTCanAct & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, ccodage, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran  from Trandiaria A where substring(A.cCodCta,6,1) = '" + psMoneda + "' and ccodope in ('" & gsCTSOTRetEfe & "','" & gsCTSOTDepEfe & "','" & gsCTSOTDepChq & "','" & gsCTSOTRetInt & "','" & gsCTSOTCanAct & "') and cFlag is null and cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If Not RSVacio(RegTrandiaria) Then
'
'        pRich = pRich + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'        pRich = pRich + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Operacion;30;Deposito;20;Retiro;20;Usu.R;6;Hora;12;", pnItem)
'
'        lsGru = ""
'
'        While Not RegTrandiaria.EOF
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                pRich = pRich + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";67;" & Format(lnTRetAge, "###,##0.00") & ";20; ;18;", pnItem)
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If lsGru <> Trim(RegTrandiaria!cCodAge) Then
'                lsGru = Trim(RegTrandiaria!cCodAge)
'                pRich = pRich + Encabezado("Agencia: " & GetNomAge(lsGru) & ";32;", pnItem)
'                lbBan = True
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnTRetAge = lnTRetAge + (RegTrandiaria!nMonTran * -1)
'
'                If RegTrandiaria!cCodOpe = gsCTSOTCanAct Then
'                    lnTRet = lnTRet + GetMonIntCanOT(RegTrandiaria!cCodCta)
'                    lnTRetAge = lnTRetAge + GetMonIntCanOT(RegTrandiaria!cCodCta)
'                End If
'            End If
'
'            pRich = pRich & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(5) & Trim(GetNomOpe(RegTrandiaria!cCodOpe)) _
'                               & Space(25 - Len(Trim(GetNomOpe(RegTrandiaria!cCodOpe)))) & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(20 - Len(lsRet)) & lsRet _
'                               & Space(6 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                               & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & oImpresora.gPrnSaltoLinea
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRich = pRich & oImpresora.gPrnSaltoPagina
'                pRich = pRich + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'                pRich = pRich + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Operacion;30;Deposito;20;Retiro;20;Usu.R;6;Hora;12;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        If lbBan Then
'            pRich = pRich + Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";67;" & Format(lnTRetAge, "###,##0.00") & ";20; ;18;", pnItem)
'        End If
'
'        pRich = pRich + Encabezado("Total Resumen;15;" & Format(lnTDep, "###,##0.00") & ";67;" & Format(lnTRet, "###,##0.00") & ";20; ;18;", pnItem)
''        pRich = pRich & Space(52 - Len(Format(lnTDep, "###,##0.00"))) & Format(lnTDep, "###,##0.00") _
'                               & Space(20 - Len(Format(lnTRet, "###,##0.00"))) & Format(lnTRet, "###,##0.00") & oImpresora.gPrnSaltoLinea
'
'        pRich = pRich & oImpresora.gPrnSaltoPagina
'        pnItem = pnItem + 1
'
'        RegTrandiaria.Close
'        Set RegTrandiaria = Nothing
'    End If
'
'    Set RegTrandiaria = Nothing
'
'End Function
'
'Public Function ReporteAperturasCTS(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim I As Long
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnValor As Double
'    Dim lnPorRet As Double
'
'    'and cCodUsu = '" & psCodUser & "'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntCTS from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSApeEfe & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntCTS from TrandiariaConsol A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSApeEfe & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldCnt, nTasaIntCTS from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSApeEfe & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
'        Exit Function
'    End If
'
'    pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'    pRTF = pRTF + Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usu.;10;Hora;17;Sald.Disp;22;TInt;11;", pnItem)
'
'    lnPorRet = (ReadParametros("23110") / 100)
'
'    While Not RegTrandiaria.EOF
'
'        If IsNull(RegTrandiaria!cNumDoc) Then
'            lsDoc = " "
'        Else
'            lsDoc = RegTrandiaria!cNumDoc
'        End If
'
'        If RegTrandiaria!nMonTran >= 0 Then
'            lnTDep = lnTDep + RegTrandiaria!nMonTran
'        Else
'            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'        End If
'
'        If IsNull(RegTrandiaria!nSaldCnt) Then
'            lnValor = 0
'        Else
'            lnValor = RegTrandiaria!nSaldCnt
'        End If
'
'        pRTF = pRTF & Space(16 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                           & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                           & Space(22 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
'                           & Space(10 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                           & Space(17 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                           & Space(22 - Len(Format(lnValor * lnPorRet, "###,##0.00"))) & Format(lnValor * lnPorRet, "###,##0.00") _
'                           & Space(11 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'
'        pnItem = pnItem + 1
'
'        If pnItem = lnRango Then
'            pRTF = pRTF & oImpresora.gPrnSaltoPagina
'            pRTF = pRTF + CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'            pRTF = pRTF + Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usu.;10;Hora;17;TInt;11;", pnItem)
'        End If
'
'        RegTrandiaria.MoveNext
'    Wend
'
'    If lnTRet <> 0 Or lnTDep <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
'
'        pnTotRet = pnTotRet + lnTRet
'        pnTotDep = pnTotDep + lnTDep
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'End Function
'
'Public Function MovimientoEfectivoCTS(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'    Dim lnValor As Double
'    Dim lnOpe As Double
'    Dim lnPorRet  As Double
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepEfe & "','" & gsCTSRetEfe & "','" & gsCTSDepNA & "','" & gsCTSRetNC & "','" & gsCTSRetInt & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from TrandiariaConsol A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepEfe & "','" & gsCTSRetEfe & "','" & gsCTSDepNA & "','" & gsCTSRetNC & "','" & gsCTSRetInt & "') and cFlag is null And DateDiff(dd,dFecTRan,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepEfe & "','" & gsCTSRetEfe & "','" & gsCTSDepNA & "','" & gsCTSRetNC & "','" & gsCTSRetInt & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RegTrandiaria.BOF And RegTrandiaria.EOF Then
'
'    Else
'
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
'
'        lnPorRet = (ReadParametros("23110") / 100)
'
'        While Not RegTrandiaria.EOF
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnOpe = RegTrandiaria!nMonTran
'            End If
'
'            If IsNull(RegTrandiaria!nSaldCnt) Then
'                lnValor = 0
'            Else
'                lnValor = RegTrandiaria!nSaldCnt
'            End If
'
'
'            If Trim(RegTrandiaria!cCodOpe) = gsCTSRetInt Then
'                pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                                   & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                                   & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
'                                   & Space(20 - Len(lsDep)) & lsDep _
'                                   & Space(20 - Len(lsRet)) & lsRet _
'                                   & Space(6 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                                   & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                                   & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'
'            Else
'                pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                                   & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                                   & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
'                                   & Space(20 - Len(lsDep)) & lsDep _
'                                   & Space(20 - Len(lsRet)) & lsRet _
'                                   & Space(6 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                                   & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                                   & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'            End If
'
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;SalAnt;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        'pRtf = pRtf + Encabezado("Resumen;15;" & Format(lnTDep, "###,##0.00") & ";57;" & Format(lnTRet, "###,##0.00") & ";20; ;25;", pnItem)
'
'        lnTRetT = lnTRet
'        lnTDepT = lnTDep
'
'        RegTrandiaria.Close
'    End If
'
'    Set RegTrandiaria = Nothing
'
'    If lnTRetT <> 0 Or lnTDepT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Format(lnTDepT, "###,##0.00") & ";57;" & Format(lnTRetT, "###,##0.00") & ";20; ;26;", pnItem)
'        pnTotRet = pnTotRet + lnTRetT
'        pnTotDep = pnTotDep + lnTDepT
'
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'Public Function MovimientoChequeCTS(pRTF As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'    Dim lnValor As Double
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnTDepT As Double
'    Dim lnTRetT As Double
'
'    Dim I As Long
'
'    lnTRet = 0
'    lnTDep = 0
'
'    If psCodUser = "XXX" Then
'        If psFecha = "" Then
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepChq & "') and cFlag is null order by A.cCodCta"
'        Else
'            SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from TrandiariaConsol A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepChq & "') and cFlag is null And DateDiff(dd,dFecTran,'" & psFecha & "') = 0 order by A.cCodCta"
'        End If
'
'    Else
'        SQL1 = "Select dFecTran, A.cCodUsu, cCodOpe, A.cCodCta, cNumDoc, nMonTran, nSaldDispCTS, nTasaIntCTS, nSaldCnt from Trandiaria A, CTS B where substring(A.cCodCta,6,1) = '" + psMoneda + "' and A.CCodCta = B.cCodCta and ccodope in ('" & gsCTSDepChq & "') and cFlag is null and A.cCodUsu = '" & psCodUser & "' order by A.cCodCta"
'    End If
'
'    RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If Not RSVacio(RegTrandiaria) Then
'
'        pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'        pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant.;20;Deposito;20;Usu.;8;Hora;14;Sald.Cnt.;20;TInt;8;", pnItem)
'
'        While Not RegTrandiaria.EOF
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lnValor = 0
'            Else
'                lnValor = RegTrandiaria!nSaldCnt
'            End If
'
'            pRTF = pRTF & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
'                               & Space(20 - Len(Format(lnValor - RegTrandiaria!nMonTran, "###,##0.00"))) & Format(lnValor - RegTrandiaria!nMonTran, "###,##0.00") _
'                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
'                               & Space(20 - Len(lsDep)) & lsDep _
'                               & Space(8 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
'                               & Space(14 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
'                               & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
'                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
'
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                pRTF = pRTF & oImpresora.gPrnSaltoPagina
'                pRTF = pRTF + CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                pRTF = pRTF + Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant.;20;Deposito;20;Usu.;8;Hora;14;Sald.Cnt.;20;TInt;8;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        lnTDepT = lnTDep
'        lnTRetT = lnTRet
'
'        pnItem = pnItem + 1
'
'        RegTrandiaria.Close
'    End If
'    Set RegTrandiaria = Nothing
'
'    If lnTDepT <> 0 Or lnTRetT <> 0 Then
'        pRTF = pRTF + Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";57;" & Trim(Format(lnTRetT, "###,##0.00")) & ";20; ;30;", pnItem)
'        pnTotDep = lnTDepT
'        pRTF = pRTF & oImpresora.gPrnSaltoPagina
'    End If
'
'End Function
'
'
