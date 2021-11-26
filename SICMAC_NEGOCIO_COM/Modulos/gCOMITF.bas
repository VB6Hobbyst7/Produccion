Attribute VB_Name = "gCOMItf"
Option Explicit

'*** Modulo para el ITF
Global gnITFPorcent As Double
Global gbITFAplica As Boolean
Global gbITFAsumidoAho As Boolean
Global gbITFAsumidoPF As Boolean
Global gbITFAsumidocreditos As Boolean
Global gbITFAsumidoGiros As Boolean

Global gnITFMontoMin As Double
Global gnITFNumTranOrigen As Long
Global gnITFNumTran As Long

'Tipo Exoneracion
Global Const gnITFTpoSinExoneracion As String = 0
Global Const gnITFTpoExoPlanilla As String = 3
Global Const gnITFTpoExoUniColegios As String = 2
Global Const gnITFTpoExoIntPublicas As String = 1
Global Const gnITFTpoExoIntFinanc As String = 6
Global Const gsRUCCmac As String = "20104888934"

Global Const gnITFTpoOpeVarias As String = "1"
Global Const gnITFTpoOpeCaja As String = "2"

Public gTCPonderadoSBS As Currency

Public Function fgITFVerificaExoneracion(ByVal psCodCta As String) As Boolean
    Dim sql As String
    Dim oCon As DCOMConecta
    Set oCon = New DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "select cCtaCod from dbo.ITFExoneracionCta Where cCtaCod = '" & psCodCta & "' And nExoTpo <> 0"
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgITFVerificaExoneracion = True
    Else
        fgITFVerificaExoneracion = False
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function


'Sin Exoneracion              0
'Planillas                    1
'Colegios/Universidades       2
'Instituciones Publicas       3
'Instituciones Financieras    4
Public Function fgITFTipoExoneracion(ByVal psCodCta As String) As Integer
    Dim sql As String
    Dim oCon As DCOMConecta
    Set oCon = New DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "select nExoTpo from dbo.ITFExoneracionCta Where cCtaCod = '" & psCodCta & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgITFTipoExoneracion = rs.Fields(0)
    Else
        fgITFTipoExoneracion = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function

'*** Obtiene los parametros de ITF
Public Sub fgITFParametros()
Dim oCon As DCOMConecta
Set oCon = New DCOMConecta
Dim lsSQL As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset
    
    lsSQL = "select nParCod, nParValor FROM PARAMETRO WHERE nParProd = 1000 And nParCod In (1001,1002,1003)"
    oCon.AbreConexion
    lr.CursorLocation = adUseClient
    Set lr = oCon.CargaRecordSet(lsSQL)
    lr.ActiveConnection = Nothing
    Do While Not lr.EOF
        Select Case lr!nParCod
            Case 1001
                gbITFAplica = IIf(lr!nParValor = 0, False, True)
            Case 1003
                gnITFPorcent = lr!nParValor
            Case 1002
                gnITFMontoMin = lr!nParValor
        End Select
        lr.MoveNext
    Loop
'    lr.Close
'
'    lsSQL = " Select cProducto, bAsumido from itfagenciaproducto where cAgeCod = '" & gsCodAge & "'"
'    lr.CursorLocation = adUseClient
'    Set lr = oCon.CargaRecordSet(lsSQL)
'    lr.ActiveConnection = Nothing
'    Do While Not lr.EOF
'        Select Case lr!cProducto
'            Case gCapAhorros
'                gbITFAsumidoAho = lr!bAsumido
'            Case gCapPlazoFijo
'                gbITFAsumidoPF = lr!bAsumido
'            Case Else
'                gbITFAsumidocreditos = lr!bAsumido
'        End Select
'        lr.MoveNext
'    Loop
    oCon.CierraConexion
    Set oCon = Nothing
    lr.Close
    Set lr = Nothing
End Sub

'*** Obtiene los parametros de ITF
Public Sub fgITFParamAsume(psAgeCod As String, Optional psProducto As String)
Dim oCon As DCOMConecta
Set oCon = New DCOMConecta
Dim lsSQL As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset
oCon.AbreConexion

gbITFAsumidoAho = False
gbITFAsumidoPF = False
gbITFAsumidoGiros = False
gbITFAsumidocreditos = False
    
lsSQL = " Select cProducto, bAsumido from itfagenciaproducto where cAgeCod = '" & psAgeCod & "' and cproducto='" & psProducto & "'"
lr.CursorLocation = adUseClient
Set lr = oCon.CargaRecordSet(lsSQL)
lr.ActiveConnection = Nothing
Do While Not lr.EOF
    Select Case lr!cProducto
        Case gCapAhorros
            gbITFAsumidoAho = lr!bAsumido
        Case gCapPlazoFijo
            gbITFAsumidoPF = lr!bAsumido
        
        Case gGiro
            gbITFAsumidoGiros = lr!bAsumido
        Case Else
            gbITFAsumidocreditos = lr!bAsumido
            
    End Select
    lr.MoveNext
Loop
oCon.CierraConexion
Set oCon = Nothing
lr.Close
Set lr = Nothing
End Sub

'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuesto(ByVal pnMonto As Double) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        
        lnValor = pnMonto * gnITFPorcent
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) > 0 Then
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        Else
            aux = CDbl(CStr(Int(lnValor)))
        End If
        lnValor = aux

        lnValor = Format(lnValor, "#0.00")
               
        
    End If
End If
fgITFCalculaImpuesto = lnValor
End Function

Public Function fgITFDesembolso(ByVal pnMonto As Double) As Double
    Dim sCadena As Currency
        sCadena = Round(pnMonto * gnITFPorcent, 6)
        fgITFDesembolso = CortaDosITF(sCadena)
End Function
Public Function CortaDosITF(ByVal lnITF As Double) As Double
Dim intpos  As Integer
Dim lnDecimal As Double
Dim lsDec As String
Dim lnEntero As Long
Dim lnPos As Long

lnEntero = Int(lnITF)
lnDecimal = Round(lnITF - Int(lnEntero), 6)
lnPos = InStr(1, Trim(Str(lnDecimal)), ".")
If lnPos > 0 Then
    lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
    lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
    lnDecimal = Val(lsDec) / 100
    CortaDosITF = lnEntero + lnDecimal
Else
    lnDecimal = 0
    CortaDosITF = lnEntero
End If
End Function
'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuestoIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean = False) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        If bCancelacion = True Then
            lnValor = Format(pnMonto / (1 + gnITFPorcent), "#0.00")
        Else
            lnValor = pnMonto / (1 + gnITFPorcent)
        End If
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
           aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
            lnValor = aux
        End If

        lnValor = Format(lnValor, "#0.00")
      
      
        'lnValor = Format(lnValor, "#0.00")
    End If
End If
fgITFCalculaImpuestoIncluido = lnValor
End Function

'*** Devuelve el Monto con el ITF agregado
Public Function fgITFCalculaImpuestoNOIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
        If bCancelacion = True Then
            lnValor = Format(pnMonto * (1 + gnITFPorcent), "#0.00")
        Else
            lnValor = pnMonto * (1 + gnITFPorcent)
        End If
        
        
        
        Dim aux As Double
        If bCancelacion = True Then
            If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
                aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
            Else
                aux = lnValor
            End If
        Else
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        End If
        
        lnValor = aux

        lnValor = Format(lnValor, "#0.00")
      
       
        'lnValor = Format(lnValor, "#0.00")
End If
fgITFCalculaImpuestoNOIncluido = lnValor
End Function


'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACRetiroImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lgsCodUser As String, oConexion As ADODB.Connection, ByVal pgsCodAge As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim lpnNroTransaccion As Long
    Dim lpsUsuRem As String
    Dim lpbGrabaTranDiaria As Integer
    Dim lbValidaMinimo As Integer
    Dim lpgsCodAge As String
   
    lbValidaMinimo = 1
    lpbGrabaTranDiaria = 1
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "spACRetiroEfectivo"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "spACRetiroEfectivo"
    Set prm = cmd.CreateParameter("psCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnMonto", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("NroDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsACRetEfe", adVarChar, adParamInput, 6)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pdFecha", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("gsCodUser", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pbGrabaTranDiaria", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bValidaMinimo", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnNroTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.spACRetiroEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, pgsCodAge, lpdFecha, lgsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACRetiroImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function


'*** Efectivo del MontoTotal ITF
Public Function fgITFEfectivoImpuesto(lcCodOpe As String, lcCodCta As String, lcNumDoc As String, lnMonTran As Currency, _
                                     lnSaldCnt As Currency, lcCodUsuRem As String, lcCodAge As String, lnTipCambio As Currency, oConexion As ADODB.Connection, ByVal pgdFecSis As Date, ByVal pgsCodUser As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter

    Dim lnTransaccion As Long
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertTranDiariaOut"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "InsertTranDiariaOut"
    Set prm = cmd.CreateParameter("dFecTran", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodUsu", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodOpe", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cNumDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nMonTran", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nSaldCnt", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nTipCambio", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("lnTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.InsertTranDiariaOut FechaHora(pgdFecSis), pgsCodUser, lcCodOpe, lcCodCta, lcNumDoc, lnMonTran, lnSaldCnt, lcCodUsuRem, lcCodAge, lnTipCambio, lnTransaccion

    fgITFEfectivoImpuesto = lnTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function

'*** Efectivo del MontoTotal ITF
Public Sub fgITFDetalleInserta(pnNumTran As Long, pcAgeCod As String, pnMonto As Currency, pnImpuesto As Currency, pnNumTranITF As Long, oConexion As ADODB.Connection)
    Dim sql As String
    sql = " Insert ITFDetalle(nNumTran, cAgeCod, nMonto, nImpuesto, nNumTranITF)" _
        & " Values (" & pnNumTran & ",'" & pcAgeCod & "'," & pnMonto & "," & pnImpuesto & "," & pnNumTranITF & ")"
    oConexion.Execute sql
End Sub

Public Sub fgITFImprimeBoleta(ByVal sNomCli As String, ByVal nMonITF As Double, ByVal psDesOpeOrigen As String, _
            Optional ByVal nNumTran As Long = 1, Optional ByVal sTitBol As String = "IMP. TRANSAC. FINANCIERAS", _
            Optional ByVal pnTipoPago As Integer = 1, Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", _
            Optional sGlosa As String, Optional pCancelaImp As Boolean = True, Optional ByVal sImp As String = "", Optional ByVal sNomAg As String = "")
'
'If pCancelaImp Then
'    Exit Sub 'cancela impresion de itf
'End If
'
'If sImp <> "" Then
'    sLpt = sImp
'End If
'
'Dim psTit As String
'Dim sMonto As String
'Dim psDesOpe As String
'Dim nFicSal As Integer
'Dim sFecha As String
'Dim sHora As String
'Dim sSep As Integer
'Dim sIni As Integer
'Dim sMax  As Integer
'Dim saux As Integer
'Dim lsNegritaOn As String
'Dim lsNegritaOff As String
'Dim lnTope As Integer
'Dim lsNroExt As String
'Dim lnCliAux As Integer
'Dim lsNomAge As String
'
'Dim lnNumLinCmac As Integer
'
'Dim lsMensaje As String * 39
'Dim lsCliAux1 As String
'Dim lsCliAux2 As String
'Dim psTexto As String
'
'Dim psTextoTC1 As String
'Dim psTextoTC2 As String
'Dim psMontoSoles As String
'Dim sPonderadoSoles As String
'
''Moneda y Cuenta porque en algunos casos no se manda la moneda... JHVP
'If Len(sCodCta) = 18 Then
'    nMonedaITF = Mid(sCodCta, 9, 1)
'End If
'
'If Len(gsNomAge) > 15 Then
'    gsNomAge = Mid(gsNomAge, 1, 15)
'End If
'
'If nMonedaITF = 1 Then
'    If pnTipoPago = 1 Then
'        psTexto = "ITF Efectivo"
'    ElseIf pnTipoPago = 2 Then
'        psTexto = "ITF Cargo a Cuenta"
'    End If
'ElseIf nMonedaITF = 2 Then
'    If pnTipoPago = 1 Then
'        psTexto = "ITF Efectivo Dolares"
'    ElseIf pnTipoPago = 2 Then
'        psTexto = "ITF Cargo a Cuenta Dolares"
'    End If
'End If
'
'lsNroExt = Str(nNumTran)
'
'ETIQ:
'
'On Error GoTo Error
'
'lnTope = 0 '6 'Tope de lineas en Boleta
'
'lsNegritaOn = gPrnNegritaON
'lsNegritaOff = gPrnNegritaOFF
'
'nFicSal = FreeFile
''If bsImp = False Then
'Open sLpt For Output As nFicSal
'
''End If
'
'Print #nFicSal, gPrnInicializa;
'
''Print #nFicSal, gPrnEspaLinea6;   'espaciamiento lineas 1/6 pulg.
''Print #nFicSal, gPrnTamPagina22;  'Longitud de página a 22 líneas'
''Print #nFicSal, gPrnTamLetra10CPI;    'Tamaño 10 cpi
''Print #nFicSal, gPrnTpoLetraRoman;       'Tipo de Letra Sans Serif
''Print #nFicSal, gPrnCondensadaOFF ' cancela condensada
''Print #nFicSal, gPrnNegritaOFF ' desactiva negrita
'
'sSep = 15
'sIni = 1
'sMax = 33
'saux = 5
'
'
'sFecha = Format$(gdFecSis, "dd/mm/yyyy")
'sHora = Format$(Time, "hh:mm:ss")
'sMonto = Format$(nMonITF, "#,##0.00")
'
'
'lsNomAge = gsNomAge
'
'If sNomAg <> "" Then
'lsNomAge = sNomAg
'End If
'
'If Len(lsNomAge) > 15 Then
'    lsNomAge = Mid(lsNomAge, 1, 15)
'End If
'
''Print #nFicSal, gPrnSaltoLinea;
''Print #nFicSal, ""
'Print #nFicSal, ""
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni); lsNegritaOn & "CMACICA"; Space(26 + sSep + saux); "CMACICA" & lsNegritaOff
'
'If nMonedaITF = 1 Then
'    Print #nFicSal, Tab(sIni); Trim(lsNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(lsNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
'Else
'    Print #nFicSal, Tab(sIni); Trim(lsNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(lsNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
'End If
'
'Print #nFicSal, ""
'
'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
'lnNumLinCmac = 0
'
'Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
'
'lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)
'
'If lnCliAux = 0 Then
'    If saux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + saux)
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(sNomCli); Space(saux + sMax + sSep - Len(sNomCli)); ImpreCarEsp(sNomCli)
'Else
'    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
'    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))
'
'    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
'    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
'
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
'
'    lnCliAux = 1
'End If
'
'If Len(Trim(sCodCta)) > 0 Then
'    sCodCta = Right("                       " & sCodCta, 18)
'    Print #nFicSal, Tab(sIni); "Cuenta:" & sCodCta; Space(8 + sSep + saux); "Cuenta:" & sCodCta
'Else
'    Print #nFicSal, ""
'End If
'
'psTit = Trim(sTitBol)
'psTit = CentrarCadena(psTit, 28)
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
'
'Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto
'
'If nMonedaITF = 1 Then
'    Print #nFicSal, ""  'JVP1
'ElseIf nMonedaITF = 2 Then
'    psTextoTC1 = "** Tipo de Cambio SBS"
'    If pnTipoPago = 1 Then
'        psTextoTC2 = "** ITF Efectivo Soles"
'    ElseIf pnTipoPago = 2 Then
'        psTextoTC2 = "** ITF Cargo a Cuenta Soles"
'    End If
'
'    If gTCPonderadoSBS = 0 Then
'        psMontoSoles = "0.00"
'        sPonderadoSoles = "0.00"
'    Else
'        psMontoSoles = Format(nMonITF * gTCPonderadoSBS, "0.00")
'        sPonderadoSoles = Format(gTCPonderadoSBS, "0.0000")
'    End If
'
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles
'
'    lnCliAux = lnCliAux + 1
'End If
'
'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita 'JVP2 JVP3
'
'lsMensaje = "Operacion Origen:"
'Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
'lsMensaje = psDesOpeOrigen
'Print #nFicSal, Tab(sIni); ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje) 'JVP4
'
'Print #nFicSal, "" 'JVP5 JVP6
'
'Print #nFicSal, sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa ''JVP6 JVP7
'
'lnTope = 7 - lnCliAux 'JVP 3-0=3 3-1=2
'
'Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
'Print #nFicSal, Tab(sIni); ImpreCarEsp(gsCodUser); Space(29 + sSep + saux); ImpreCarEsp(gsCodUser)
'
'Dim clsGen As DCOMGeneral
'Set clsGen = New DCOMGeneral
'lsMensaje = clsGen.GetMensajeBoletas(sCodCta)
'Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
'
'lnNumLinCmac = lnNumLinCmac + 1
'
'For saux = 1 To (lnTope - lnNumLinCmac) 'JVP DE 1 A 3-1=2 DE 1 A 2-1=1
'    Print #nFicSal, ""
'Next saux
'Close nFicSal
'Exit Sub
'Error:
'    Close nFicSal
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If

End Sub

Public Function fgITFGetTitular(psCtaCod As String) As String
    Dim sql As String
    Dim oCon As DCOMConecta
    Set oCon = New DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    sql = " Select Top 1 dbo.PstaNombre(cPersNombre,1) Nombre From ProductoPersona PP" _
        & " Inner Join Persona PE ON PP.cPersCod = PE.cPersCod" _
        & " Where PP.cCtaCod = '" & psCtaCod & "' Order By cPersNombre"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        fgITFGetTitular = ""
    Else
        fgITFGetTitular = rs!NOMBRE
    End If
    
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function fgITFGetNumtranOrigen(oConexion As ADODB.Connection) As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " SELECT @@IDENTITY "
    rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        fgITFGetNumtranOrigen = 0
    Else
        fgITFGetNumtranOrigen = rs.Fields(0)
    End If
    
End Function

'Public Sub fgITFExtornaITF_Trandiaria(psCtaCod As String, pCo As ADODB.Connection)
'    Dim sql As String
'
'    sql = "Update Trandiaria Set cFlag='X' Where cOpeCod='" & gcITFOpePrend & "' and cCodCta='" & psCtaCod & "'"
'    pCo.Execute (sql)
'End Sub

'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACDepositoImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lgsCodUser As String, oConexion As ADODB.Connection, ByVal pgsCodAge As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim lpnNroTransaccion As Long
    Dim lpsUsuRem As String
    Dim lpbGrabaTranDiaria As Integer
    Dim lbValidaMinimo As Integer
    Dim lpgsCodAge As String
   
    lbValidaMinimo = 1
    lpbGrabaTranDiaria = 1
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "spACDepositoEfectivo"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "spACDepositoEfectivo"
    Set prm = cmd.CreateParameter("psCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnMonto", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("NroDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsACRetEfe", adVarChar, adParamInput, 6)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pdFecha", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("gsCodUser", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pbGrabaTranDiaria", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bValidaMinimo", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnNroTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.spACDepositoEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, pgsCodAge, lpdFecha, lgsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACDepositoImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function


Public Function GetnNumtranRemoto(oConexion As ADODB.Connection, psFecha As String, psCtaCod As String, Optional psOpeCod As String = "", Optional ByVal pbExtornoOtroDia As Boolean = False) As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If psOpeCod = "" Then
        If Not pbExtornoOtroDia Then
            sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From Trandiaria " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "')"
        Else
            sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From TrandiariaConsol " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "')"

        End If
    Else
        If Not pbExtornoOtroDia Then
            sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From Trandiaria " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "' and cCodOpe = '" & psOpeCod & "')"
        Else
            sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From TrandiariaConsol " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "' and cCodOpe = '" & psOpeCod & "')"
        End If
    End If
    rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        GetnNumtranRemoto = 0
    Else
        If IsNull(rs!nNumTran) Then
            GetnNumtranRemoto = 0
        Else
            GetnNumtranRemoto = rs!nNumTran
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function


Public Function VerifOpeVariasAfectaITF(psOpeCod As String) As Boolean
    Dim sql As String
    Dim oCon As DCOMConecta
    Set oCon = New DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "Select cOpeCod From ITFOperaciones Where nTipo = 1 And cOpeCod = '" & Trim(psOpeCod) & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        VerifOpeVariasAfectaITF = False
    Else
        VerifOpeVariasAfectaITF = True
    End If
End Function



