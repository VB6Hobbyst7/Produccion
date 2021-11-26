Attribute VB_Name = "gITF"
Option Explicit

'*** Modulo para el ITF
Global gnITFPorcent As Double
Global gbITFAplica As Boolean
Global gnITFMontoMin As Double
Global gnITFNumTranOrigen As Long
Global gnITFNumTran As Long

Global Const gcITFOpePrend As String = "030120"
Global Const gsITFACCargoCuenta As String = "200320"
Global Const gsITFACCargoCuentaExt As String = "201320"
Global Const gsITFEfectivo As String = "030120"
Global Const gsITFEfectivoExt As String = "030199"

'Tipo Exoneracion
Global Const gnITFTpoSinExoneracion As String = 0
Global Const gnITFTpoExoPlanilla As String = 1
Global Const gnITFTpoExoUniColegios As String = 2
Global Const gnITFTpoExoIntPublicas As String = 3
Global Const gnITFTpoExoIntFinanc As String = 4

'Public Sub InsertaExonerado(pCo As DConecta, pcCodCta As String, _
'                              pTipo As String, pFechaHora As String, pUser As String)
'
'Dim rs As New ADODB.Recordset
'Dim sql As String
'Dim cEst As String
'sql = "Select * from ITFCtaExonerada Where cCodCta = '" & pcCodCta & "'"
'
'Set rs = pCo.Execute(sql)
'If rs.EOF And rs.BOF Then
'    sql = " Insert ITFCtaExonerada "
'    sql = sql & "(cCodCta,cTipo,dFechaReg,cUser) values"
'    sql = sql & "('" & pcCodCta & "','" & pTipo & "','" & pFechaHora & "','" & pUser & "')"
'    pCo.Execute (sql)
'    cEst = "E"
'    MsgBox "La cta. " & pcCodCta & " se exoneró del ITF", vbInformation, "AVISO"
'Else
'    If IsNull(rs!cFlag) Then
'        sql = "Update ITFCtaExonerada Set cFlag='X' where ='" & pcCodCta & "'"
'        cEst = "X"
'        MsgBox "La cta. " & pcCodCta & " se extorno", vbInformation, "AVISO"
'    Else
'        sql = "Update ITFCtaExonerada Set cFlag=Null, dFecha='" & pFechaHora & "' where ='" & pcCodCta & "'"
'        cEst = "E"
'        MsgBox "La cta. " & pcCodCta & " se exoneró del ITF", vbInformation, "AVISO"
'    End If
'End If
'sql = "Insert ITFCtaExoneradaDet "
'sql = sql & " (dFecha,cCodCta,cUser,cEstado) values"
'sql = sql & " ('" & pFechaHora & "','" & pcCodCta & "','" & pUser & "','" & cEst & "')"
'pCo.Execute (sql)
'End Sub
   
Public Function VerificaExoneracion(ByVal psCodCta As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psCodCta = "" Then
        VerificaExoneracion = False
        Exit Function
    End If
    oCon.AbreConexionRemota Left(psCodCta, 2)
    
    sql = "select cCodCta from dbo.ITFCtaExonerada where cCodCta='" & psCodCta & "' and  cFlag is null"
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.EOF) Then
        VerificaExoneracion = True
    Else
        VerificaExoneracion = False
    End If
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Sin Exoneracion              0
'Planillas                    1
'Colegios/Universidades       2
'Instituciones Publicas       3
'Instituciones Financieras    4
Public Function fgTipoExoneracion(ByVal psCodCta As String) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexionRemota Left(psCodCta, 2)
    
    sql = "select cTipo from dbo.ITFCtaExonerada where cCodCta = '" & psCodCta & "' and  cFlag is null"
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgTipoExoneracion = rs!cTipo
    Else
        fgTipoExoneracion = 0
    End If
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** Obtiene los parametros de ITF
Public Sub fgITFParametros()
    Dim lsSQL As String
    Dim lr As ADODB.Recordset
    Set lr = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexionRemota "01"
    lsSQL = "Select * from Parametro where cCodPar like '5%' "
    Set lr = oCon.CargaRecordSet(lsSQL)
    Do While Not lr.EOF
        Select Case lr!cCodPar
            Case "50001"
                gbITFAplica = IIf(lr!nValor1 = 0, False, True)
            Case "50002"
                gnITFPorcent = lr!nValor1
            Case "50003"
                gnITFMontoMin = lr!nValor1
        End Select
        lr.MoveNext
    Loop
    lr.Close
    Set lr = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuesto(ByVal pnMonto As Double) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        lnValor = pnMonto * gnITFPorcent
        lnValor = Round(lnValor, 2)
    End If
End If
fgITFCalculaImpuesto = lnValor
End Function

'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuestoIncluido(ByVal pnMonto As Double) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        lnValor = pnMonto / (1 + gnITFPorcent)
        lnValor = Round(lnValor, 2)
    End If
End If
fgITFCalculaImpuestoIncluido = lnValor
End Function

'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACRetiroImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lgsCodUser As String, oConexion As ADODB.Connection) As Long
    
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
    oConexion.spACRetiroEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, gsCodAge, lpdFecha, lgsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACRetiroImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function


'*** Efectivo del MontoTotal ITF
Public Function fgITFEfectivoImpuesto(lcCodOpe As String, lcCodCta As String, lcNumDoc As String, lnMonTran As Currency, _
                                     lnSaldCnt As Currency, lcCodUsuRem As String, lcCodAge As String, lnTipCambio As Currency, oConexion As ADODB.Connection) As Long
    
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
    oConexion.InsertTranDiariaOut FechaHora(gdFecSis), gsCodUser, lcCodOpe, lcCodCta, lcNumDoc, lnMonTran, lnSaldCnt, lcCodUsuRem, lcCodAge, lnTipCambio, lnTransaccion

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

Public Sub fgITFImprimeBoleta(ByVal sNomCli As String, ByVal nMonITF As Double, psDesOpeOrigen As String, Optional nNumTran As Long = 1, Optional sTitBol As String = "IMP. TRANSAC. FINANCIERAS", Optional pnTipoPago As Integer = 1, Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", Optional sGlosa As String)
Dim psTit As String
Dim sMonto As String
Dim psDesOpe As String
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax  As Integer
Dim sAux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lnTope As Integer
Dim lsNroExt As String
Dim lnCliAux As Integer
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39
Dim lsCliAux1 As String
Dim lsCliAux2 As String
Dim psTexto As String

If pnTipoPago = 1 Then
    psTexto = "Monto ITF Efectivo"
ElseIf pnTipoPago = 2 Then
    psTexto = "Monto ITF Cargo a Cuenta"
End If

lsNroExt = Str(nNumTran)


ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = gPrnNegritaON
lsNegritaOff = gPrnNegritaOFF
  
nFicSal = FreeFile
Open sLPT For Output As nFicSal

Print #nFicSal, gPrnInicializa;

sSep = 15
sIni = 1
sMax = 33
sAux = 5


sFecha = Format$(gdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(nMonITF, "#,##0.00")
 
lsNomAge = gsNomAge

'Print #nFicSal, gPrnSaltoLinea;
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT"; Space(28 + sSep + sAux); "CMACT"

If nMonedaITF = 1 Then
    Print #nFicSal, Tab(sIni); Trim(gsNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(gsNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(gsNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(gsNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If
 
Print #nFicSal, ""
 
Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
lnNumLinCmac = 0

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sAux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + sAux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(sNomCli); Space(sAux + sMax + sSep - Len(sNomCli)); ImpreCarEsp(sNomCli)
Else
    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))

    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)

    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)

    lnCliAux = 1
End If

If Len(Trim(sCodCta)) > 0 Then
    sCodCta = Right("            " & sCodCta, 12)
    Print #nFicSal, Tab(sIni); "Cuenta:" & sCodCta; Space(14 + sSep + sAux); "Cuenta:" & sCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(sTitBol)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
 
Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto
Print #nFicSal, ""
  
Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

'Print #nFicSal, ""

lsMensaje = "Operacion Origen:"
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
lsMensaje = psDesOpeOrigen
Print #nFicSal, Tab(sIni); ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje)

Print #nFicSal, ""
 
Print #nFicSal, sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa
     
lnTope = 3 - lnCliAux

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(gsCodUser); Space(29 + sSep + sAux); ImpreCarEsp(gsCodUser)

lsMensaje = ""
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1

For sAux = 1 To (lnTope - lnNumLinCmac)
    Print #nFicSal, ""
Next sAux
Close nFicSal
Exit Sub
ERROR:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "") = vbYes Then
        GoTo ETIQ
    End If

End Sub

Public Function fgITFGetTitular(psCtaCod As String, pbCreditos As Boolean, oConexion As ADODB.Connection) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If Not pbCreditos Then
        sql = " Select top 1 cNomPers from perscuenta pc" _
            & " inner join dbPersona..persona pe on pc.ccodpers = pe.ccodpers" _
            & " where ccodcta = '" & psCtaCod & "' And cRelaCta = 'TI' order by  cNomPers"
    Else
        sql = " Select top 1 cNomPers from perscredito pc" _
            & " inner join dbPersona..persona pe on pc.ccodpers = pe.ccodpers" _
            & " where ccodcta = '" & psCtaCod & "' And cRelaCta = 'TI' order by  cNomPers"
    End If
    
    rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        fgITFGetTitular = ""
    Else
        fgITFGetTitular = rs!cNomPers
    End If
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

Public Sub fgITFExtornaITF_Trandiaria(psCtaCod As String, pCo As ADODB.Connection)
    Dim sql As String

    sql = "Update Trandiaria Set cFlag='X' Where cOpeCod='" & gcITFOpePrend & "' and cCodCta='" & psCtaCod & "'"
    pCo.Execute (sql)
End Sub

'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACDepositoImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lgsCodUser As String, oConexion As ADODB.Connection) As Long
    
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
    oConexion.spACDepositoEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, gsCodAge, lpdFecha, lgsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACDepositoImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function

Public Function fgITFExtonoImpuesto(pnNumTran As Long, psCodCta As String, pnMontoITF As Currency, oConexion As ADODB.Connection, oConRemota As ADODB.Connection, Optional pbITFEfectivo As Boolean = True, Optional pbITFLocal As Boolean = True) As Integer
    Dim sql As String
    Dim lnNumTran As Long
    Dim lnMonITF As Currency
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If pbITFEfectivo Then
        If pbITFLocal Then
            sql = " Select Top 1 nImpuesto, nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ""
            rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not (rs.EOF And rs.BOF) Then
                lnNumTran = rs!nNumTranITF
                lnMonITF = rs!nImpuesto
                pnMontoITF = lnMonITF
                rs.Close
                
                sql = " Update TransAho " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                sql = " Update Trandiaria " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                If lnMonITF <> 0 Then
                    fgITFExtonoImpuesto = 1
                Else
                    fgITFExtonoImpuesto = 0
                End If
            
                gnITFNumTran = fgITFEfectivoImpuesto(gsITFEfectivoExt, psCodCta, "", lnMonITF, 0, "", gsCodAge, 0, oConexion)
            
                sql = " Update ITFDetalle Set nNumTranITFExt = " & gnITFNumTran & " Where nNumTran = " & pnNumTran & ""
                oConexion.Execute sql
            Else
                fgITFExtonoImpuesto = 0
            End If
        Else
            sql = " Select Top 1 nImpuesto, nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ""
            rs.Open sql, oConRemota, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not (rs.EOF And rs.BOF) Then
                lnNumTran = rs!nNumTranITF
                lnMonITF = rs!nImpuesto
                pnMontoITF = lnMonITF
                rs.Close
                
                sql = " Update TransAho " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                sql = " Update Trandiaria " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                If lnMonITF <> 0 Then
                    fgITFExtonoImpuesto = 1
                Else
                    fgITFExtonoImpuesto = 0
                End If
            
                gnITFNumTran = fgITFEfectivoImpuesto(gsITFEfectivoExt, psCodCta, "", lnMonITF, 0, "", gsCodAge, 0, oConexion)
            
                sql = " Update ITFDetalle Set nNumTranITFExt = " & gnITFNumTran & " Where nNumTran = " & pnNumTran & ""
                oConRemota.Execute sql
            Else
                rs.Close
                fgITFExtonoImpuesto = 0
            End If
        End If
    Else
        If pbITFLocal Then
            sql = " Select Top 1 nImpuesto, nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ""
            rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not (rs.EOF And rs.BOF) Then
                lnNumTran = rs!nNumTranITF
                lnMonITF = rs!nImpuesto
                pnMontoITF = lnMonITF
                rs.Close
                
                sql = " Update TransAho " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                sql = " Update Trandiaria " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConexion.Execute sql
                
                If lnMonITF <> 0 Then
                    fgITFExtonoImpuesto = 1
                Else
                    fgITFExtonoImpuesto = 0
                End If
                
                gnITFNumTran = fgITFACDepositoImpuesto(psCodCta, lnMonITF, "", gsITFACCargoCuentaExt, FechaHora(gdFecSis), gsCodUser, oConexion)
            
                sql = " Update ITFDetalle Set nNumTranITFExt = " & gnITFNumTran & " Where nNumTran = " & pnNumTran & ""
                oConexion.Execute sql
            Else
                rs.Close
                fgITFExtonoImpuesto = 0
            End If
        Else
            sql = " Select Top 1 nImpuesto, nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ""
            rs.Open sql, oConRemota, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not (rs.EOF And rs.BOF) Then
                lnNumTran = rs!nNumTranITF
                lnMonITF = rs!nImpuesto
                pnMontoITF = lnMonITF
                rs.Close
                
                sql = " Update TransAho " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConRemota.Execute sql
                
                sql = " Update Trandiaria " _
                    & " Set cFlag = 'X' Where nNumtran = " & lnNumTran
                oConRemota.Execute sql
                
                If lnMonITF <> 0 Then
                    fgITFExtonoImpuesto = 1
                Else
                    fgITFExtonoImpuesto = 0
                End If
                
                gnITFNumTran = fgITFACDepositoImpuesto(psCodCta, lnMonITF, "", gsITFACCargoCuentaExt, FechaHora(gdFecSis), gsCodUser, oConRemota)
            
                sql = " Update ITFDetalle Set nNumTranITFExt = " & gnITFNumTran & " Where nNumTran = " & pnNumTran & ""
                oConRemota.Execute sql
            Else
                rs.Close
                fgITFExtonoImpuesto = 0
            End If
        End If
    End If
End Function


Public Function GetnNumtranRemoto(oConexion As ADODB.Connection, psFecha As String, psCtaCod As String, Optional psOpeCod As String = "") As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If psOpeCod = "" Then
        sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
            & " Select nNumTran From Trandiaria " _
            & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
            & " and cCodCta = '" & psCtaCod & "')"
    Else
        sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
            & " Select nNumTran From Trandiaria " _
            & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
            & " and cCodCta = '" & psCtaCod & "' and cCodOpe = '" & psOpeCod & "')"
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


'Public Function fgITFOperacionConImpuesto(pnNumTran As Long, oConexion As ADODB.Connection, Optional pbITFEfectivo As Boolean = True) As Boolean
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    If pbITFEfectivo Then
'        sql = " Update TransAho " _
'            & " Set cFlag = 'X'" _
'            & " Where nNumtran = (Select Top 1 nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ")"
'        oConexion.Execute sql
'
'        sql = " Update Trandiaria " _
'            & " Set cFlag = 'X'" _
'            & " Where nNumtran = (Select Top 1 nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ")"
'        oConexion.Execute sql
'
'        sql = " Select @@ROWCOUNT"
'        rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
'
'        fgITFExtonoImpuesto = rs.Fields(0)
'
'        rs.Close
'        Set rs = Nothing
'    Else
'        sql = " Select * From Trandiaria " _
'            & " Where nNumtran = (Select Top 1 nNumTranITF From ITFDetalle Where nNumTran = " & pnNumTran & ")"
'        rs.Open sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
'
'    End If
'
'End Function



