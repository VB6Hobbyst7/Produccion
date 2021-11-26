VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapAbonoCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
            ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
            ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
            Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
            Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
            Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
            Optional pdFecSis As Date, Optional psNomAge As String = "") As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim sAux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39

ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open gsLpt For Output As nFicSal

Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
sAux = 5


sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = psNomAge

'Print #nFicSal, Chr$(10);
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"

If Mid(psCodCta, 9, 1) = 1 Then
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If

If psNomAgeRem = "" Then
    Print #nFicSal, ""
Else
    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

'psCliente = PstaNombre(psCliente)

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(sAux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(psTit)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"


lnChq = InStr(1, psText, "*", vbTextCompare)

If psTexto = "XXX" Then
    If lnChq = 0 Then
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto
        Print #nFicSal, ""
    Else
        lsChqAux1 = (Mid(psText, 1, lnChq - 1))
        lsChqAux2 = (Mid(psText, lnChq + 1))
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2))); Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2)))
    End If
Else
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(sAux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
    Print #nFicSal, ""
End If


Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

If pbSaldoInt Then
    If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
        If pbOpSaldoC Then
            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
        Else
            Print #nFicSal, ""
        End If
    Else
        Print #nFicSal, ""
        Print #nFicSal, ""
        pbSaldoInt = False
    End If
Else
    Print #nFicSal, ""
End If

lsInteres = pnInteresA

If pbSaldoInt Then
    If lsInteres <> "No Valido" Then
        lsInteres = Format(lsInteres, "#,##0.00")
        Print #nFicSal, Tab(sIni); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres; Space(-1 + sSep); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres
    End If
Else
    Print #nFicSal, ""
End If

If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

If Not psNumDias = "----" Then
    Print #nFicSal, Tab(sIni); "Nro Dias Interes"; Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)); psNumDias; Space(-1 + sSep); "Nro Dias Interes"; Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")); psNumDias
    lnTope = 4 - lnCliAux
Else
    lnTope = 3 - lnCliAux
End If

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
If psCodUsuRem = "" Then
    Print #nFicSal, Tab(sIni); ImpreCarEsp(gsCodUser); Space(29 + sSep + sAux); ImpreCarEsp(gsCodUser)
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(gsCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(gsCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1
 
'Print #nFicSal, Chr$(12)
For sAux = 1 To (lnTope - lnNumLinCmac)
    Print #nFicSal, ""
Next sAux
Close nFicSal
Exit Function
ERROR:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function


Private Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecSis As Date, Optional psNomAge As String = "") As String

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = sMsgProd
Dim clsMant As NCapMantenimiento
Set clsMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
bReImp = False
Do
    If bDocumento Then
        Select Case nDocumento
            Case TpoDocCheque
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocNotaAbono, TpoDocNotaCargo
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocOrdenPago
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
        End Select
    End If
    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bReImp = True
    Else
        bReImp = False
    End If
Loop Until Not bReImp
End Function

Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim clsMant As NCapMantenimiento
Dim clsPar As nCapDefinicion
Dim rsCta As Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))
Set clsPar = New nCapDefinicion
If nMoneda = gMonedaNacional Then
    nMontoMinimo = clsPar.GetCapParametro(gSaldMinAhoMN)
Else
    nMontoMinimo = clsPar.GetCapParametro(gSaldMinAhoME)
End If
Set clsPar = Nothing
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function

Public Function CapCargoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional bOPExiste As Boolean = False, Optional bActivaCta As Boolean = True, _
            Optional bNumExtracto As Boolean = True, Optional dFecSis As Date, _
            Optional sNomAge As String) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, I As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
Randomize
For I = 0 To Rnd(2000) * 1000
Next I
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If sGlosa = "" Then
    sGlosa = "Cargo Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & ". Cargo Cuenta = " & sCuenta
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, gdFecSis)
nSaldoCnt = nSaldoCnt - nMonto
clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto
clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            clsCap.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
      sMsgOpe = "Retiro NC"
    End If
Else
    sMsgOpe = "Retiro Efectivo"
End If
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, gdFecSis, sMovNro
End If
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro

clsCap.dbCmact.CommitTrans
CapCargoCuentaAho = nSaldoCnt
If sNroDoc = "" Then
    EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , dFecSis, sNomAge
Else
    EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, , , dFecSis, sNomAge
End If
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoCuentaAho = 0
End Function

Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional dFecSis As Date, Optional sNomAge As String = "", Optional ByVal pbUsaTransaccion As Boolean = True) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String
Dim nMovNro As Long
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, gdFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, gdFecSis)
nSaldoCnt = nSaldoCnt + nMonto
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
If sNroDoc = "" Then ' Si la operacion es en efectivo
    clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto
    clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Depósito Efectivo"
Else
    If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
        nSaldoDisp = nSaldoDisp + 0
        nIntGanado = 0
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro
        sMsgOpe = "Depósito Cheque"
    ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
        clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        sMsgOpe = "Depósito Nota Abono"
    End If
End If
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, gdFecSis, sMovNro
End If
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapAbonoCuentaAho = nSaldoCnt
If sNroDoc = "" Then
    EmiteBoleta "DEPOSITO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , dFecSis, sNomAge
Else
    EmiteBoleta "DEPOSITO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, dFechaValor, , dFecSis, sNomAge
    
End If
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function
