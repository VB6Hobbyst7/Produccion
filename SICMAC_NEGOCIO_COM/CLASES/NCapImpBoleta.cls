VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapImpBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ImprimeCertificadoPlazoFijo(ByVal dApertura As Date, ByVal sNomCli As String, _
        ByVal sDirCli As String, ByVal sCuenta As String, nExtracto As String, _
        ByVal nPlazo As Long, ByVal nSaldo As Double, ByVal nTasa As Double, _
        ByVal sRetInt As String, ByVal sLetras As String, Optional nDuplicado As Integer = 0, _
        Optional bCadena As Boolean = False) As String
        
Dim intLinBla As Integer, nFicSal As Integer
Dim sFchLet As String, sFchVct As String
Dim sFchCan As String, sForMon As String
Dim sVctLet As String, dVencimiento As Date
Dim sTaInNo As String, sTaInEf As String
Dim sTaInAn As String, sPlazo As String
Dim i As Integer
Dim sCad As String

sFchLet = ArmaFecha(dApertura)
dVencimiento = DateAdd("d", nPlazo, dApertura)
sVctLet = ArmaFecha(dVencimiento)
sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", "US$ ")

sTaInNo = Format$((nTasa / 12), "#0.00")
sTaInEf = Format$(((((((nTasa / 12) / 3000) + 1) ^ 30) - 1) * 100), "#0.00")
sTaInAn = Format(((((((nTasa / 12) / 3000) + 1) ^ 360) - 1) * 100), "#0.00")

sPlazo = Trim(nPlazo) & " dia(s)"


sCad = sCad & Chr$(15)                           'Retorna al tipo de letra normal
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
sCad = sCad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
sCad = sCad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
sCad = sCad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna

For i = 1 To 6
    sCad = sCad & Chr$(10)     '   "Fila No. " & Trim(Str(i))
Next i

'       Impresión
sCad = sCad & Space(43) & UCase(sFchLet) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & ImpreCarEsp(sNomCli) & Chr$(10)      '  psNomCli
sCad = sCad & sDirCli & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "PLAZO FIJO " & Space(24) & sCuenta & Space(17) & FillNum(Trim(nExtracto), 5, " ") & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "DEPOSITO A PLAZO FIJO " & Space(9) & ":" & Space(3) & Trim(nPlazo) & " dias " & Space(14) & Chr$(27) + Chr$(69) & sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2) & Chr$(27) + Chr$(70) & Chr$(10)
sCad = sCad & "FECHA DE VENCIMIENTO               " & Chr$(10)
sCad = sCad & "Y/O CANCELACION                :   " & UCase(sVctLet) & Chr$(10)
sCad = sCad & Chr$(10) '"TASA DE INTERES NOMINAL        :   " & sTaInNo & "%  MENSUAL" & Chr$(10)
sCad = sCad & "TASA DE INTERES EFECTIVA       :   " & sTaInEf & "%  MENSUAL " & Chr$(10)
sCad = sCad & "FRECUENCIA DE CAPITALIZACION   :   DIARIA" & Chr$(10)
sCad = sCad & "LOS INTERESES SE ABONARAN      :   " & Left(sRetInt, 20) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(27) + Chr$(69)                           'Establece tipo de letra negrita
sCad = sCad & "SON :" & Chr$(10)
sCad = sCad & sLetras
sCad = sCad & Chr$(10) 'Space(54) & sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "TASA DE INTERES EFECTIVA ANUAL :   " & sTaInAn & "%"
sCad = sCad & Chr$(27) + Chr$(70)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(27) + Chr$(33) + Chr$(40)
sCad = sCad & IIf(nDuplicado = 0, "", "DUPLICADO " & Trim(nDuplicado)) & Chr$(10)
sCad = sCad & Chr$(27) + Chr$(33) + Chr$(0)
sCad = sCad & Chr$(27) + Chr$(70)                           'Desactiva tipo de letra negrita
sCad = sCad & Chr$(12)                           'Avance de Página
If bCadena Then
    ImprimeCertificadoPlazoFijo = sCad
    Exit Function
Else
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sCad
    Close nFicSal
End If
End Function


Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional bCadena As Boolean = False, Optional bPreguntaSaldo As Boolean = True, Optional pnITFValor As Currency = 0) As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim saux As Integer
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

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)


sCad = sCad & Chr$(27) & Chr$(64)

sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

'scad = scad & Chr$(10);
sCad = sCad & lsNegritaOn 'Activa Negrita


sCad = sCad & Chr(10) & lsNegritaOn 'Activa Negrita
sCad = sCad & Space(sIni) & "CMACICA - AHORRO" & Space(19 + sSep + saux) & "CMACICA - AHORRO" & Chr$(10)

If Mid(psCodCta, 9, 1) = 1 Then
    sCad = sCad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
Else
    sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10) 'Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
End If

If psNomAgeRem = "" Then
    sCad = sCad & Chr$(10)
Else
    sCad = sCad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        sCad = sCad & lsNegritaOff 'Desactiva Negrita
    Else
        sCad = sCad & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & lsNegritaOff & psLin3 & lsNegritaOff & Chr$(10)   'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    sCad = sCad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   '"NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
    lnNumLinCmac = 1
End If

sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Chr$(10)

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
    sCad = sCad & Space(sIni) & ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & Chr$(10)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
    
    sCad = sCad & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & Chr$(10)
    sCad = sCad & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & Chr$(10)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    sCad = sCad & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & Chr$(10)
Else
    sCad = sCad & Chr$(10)
End If

psTit = Trim(psTit)
psTit = CentrarCadena(psTit, 28)
sCad = sCad & lsNegritaOn 'Activa Negrita
sCad = sCad & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr$(10)


lnChq = InStr(1, psText, "*", vbTextCompare)

If psTexto = "XXX" Then
    If lnChq = 0 Then
        sCad = sCad & Space(sIni) & ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
        sCad = sCad & Chr$(10)
    Else
        lsChqAux1 = (Mid(psText, 1, lnChq - 1))
        lsChqAux2 = (Mid(psText, lnChq + 1))
        sCad = sCad & Space(sIni) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
        sCad = sCad & Space(sIni) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
    End If
Else
    sCad = sCad & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & Chr$(10)
    sCad = sCad & Chr$(10)
End If

sCad = sCad & lsNegritaOff 'Desactiva Negrita

If pbSaldoInt Then
    If bPreguntaSaldo Then
        If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
            If pbOpSaldoC Then
                sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
            Else
                sCad = sCad & Chr$(10)
            End If
        Else
            sCad = sCad & Chr$(10)
            sCad = sCad & Chr$(10)
            pbSaldoInt = False
        End If
    Else
        sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
        If pbOpSaldoC Then
            sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
        Else
            sCad = sCad & Chr$(10)
        End If
    End If
Else
    sCad = sCad & Chr$(10)
End If

lsInteres = pnInteresA

If pbSaldoInt Then
    If lsInteres <> "No Valido" Then
        lsInteres = Format(lsInteres, "#,##0.00")
        sCad = sCad & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
    End If
Else
    sCad = sCad & Chr$(10)
End If

If Not psNumDoc = "" Then
    sCad = sCad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
Else
    sCad = sCad & Chr$(10)
End If

If Not psNumDias = "----" Then
    sCad = sCad & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & Chr$(10)
    lnTope = 5 - lnCliAux
Else
    lnTope = 4 - lnCliAux
End If

sCad = sCad & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
If psCodUsuRem = "" Then
    sCad = sCad & Space(sIni) & ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & ImpreCarEsp(sCodUser) & Chr$(10)
Else
    sCad = sCad & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Chr$(10)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
sCad = sCad & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & lsNegritaOff & ImpreCarEsp(lsMensaje) & lsNegritaOff & Chr$(10)


If pnITFValor > 0 Then
    lnNumLinCmac = lnNumLinCmac + 2
    sCad = sCad & Space(sIni) & "ITF : " & Format(pnITFValor, "#,##0.00") & Space(29 + sSep + saux) & "ITF : " & Format(pnITFValor, "#,##0.00") & Chr$(10)
Else
    lnNumLinCmac = lnNumLinCmac + 1
End If

'scad = scad & Chr$(12)
Dim nmenos As Integer
  nmenos = 0
If psCodOpe = CStr(gAhoTransCargo) Or psCodOpe = CStr(gAhoTransAbono) Then
    nmenos = 2
End If

For saux = 1 To (lnTope - lnNumLinCmac) - nmenos
    sCad = sCad & Chr$(10)
Next saux

If bCadena Then
    ImprimeBoleta = sCad
Else
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sCad
    Close nFicSal
    ImprimeBoleta = ""
End If
Exit Function
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function

Private Function LinRec(ByVal sTexto As String, pnSepara As Integer) As String
LinRec = sTexto & Space(pnSepara) & Chr$(10) ' sTexto & Chr$(10)
End Function
Public Function ImprimeBoletaLavadoDinero(ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal dFecha As Date, ByVal sCuenta As String, ByVal sTitNombre As String, _
        ByVal sTitDocId As String, ByVal sTitDireccion As String, ByVal sReaNombre As String, _
        ByVal sReaDocId As String, ByVal sReaDireccion As String, ByVal sBenNombre As String, _
        ByVal sBenDocId As String, ByVal sBenDireccion As String, ByVal sOperacion As String, _
        ByVal nImporte As Double, ByVal sLpt As String, Optional sAgeRem As String = "", _
        Optional bCadena As Boolean = False, Optional sTipo As String = "", Optional nMoneda As Integer = 1) As String
    
Dim sImpre As String, sCad As String
Dim nFicSal As Integer
Dim sBON As String, sBOFF As String, sOpe As String
Dim nFila As Integer, nSep As Integer, nMax As Integer
Dim sTexto As String

ETIQ:
    On Error GoTo ImpreErr
    

sCad = sCad & Chr$(27) & Chr$(64)

sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita
    
nSep = 14
nMax = 39

sBON = Chr$(27) & Chr$(71)
sBOFF = Chr$(27) & Chr$(72)

sCad = sCad & sBON & Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & sBOFF & Space(nSep) & sBON & Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & sBOFF & Chr$(10)
sCad = sCad & sBON & Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & sBOFF & Space(nSep) & sBON & Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & sBOFF & Chr$(10)
sCad = sCad & sBON & Justifica("TRANS.EFE. " & sBON & "(Cta: " & sCuenta & ") " & sBOFF, nMax + 2) & sBOFF & Space(nSep + 2) & sBON & Justifica("TRANS.EFE. " & sBON & "(Cta: " & sCuenta & ") " & sBOFF, nMax + 2) & sBOFF & Chr$(10)
sCad = sCad & sBON & Justifica("-------------  TITULAR  ---------------", nMax) & sBOFF & Space(nSep) & sBON & Justifica("-------------  TITULAR  ---------------", nMax) & sBOFF & Chr$(10)
sCad = sCad & Justifica("NOMBRE : " & sTitNombre, nMax) & Space(nSep) & Justifica("NOMBRE : " & sTitNombre, nMax) & Chr$(10)
sCad = sCad & Justifica("DOC ID : " & sTitDocId, nMax) & Space(nSep) & Justifica("DOC ID : " & sTitDocId, nMax) & Chr$(10)
sCad = sCad & Justifica("DIREC. : " & sTitDireccion, nMax) & Space(nSep) & Justifica("DIREC. : " & sTitDireccion, nMax) & Chr$(10)

sCad = sCad & sBON & Justifica("----  PERSONA REALIZA TRANSACCIÓN  ----", nMax) & sBOFF & Space(nSep) & sBON & Justifica("----  PERSONA REALIZA TRANSACCIÓN  ----", nMax) & sBOFF & Chr$(10)
sCad = sCad & Justifica("NOMBRE : " & sReaNombre, nMax) & Space(nSep) & Justifica("NOMBRE : " & sReaNombre, nMax) & Chr$(10)
sCad = sCad & Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & Space(nSep) & Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & Chr$(10)
sCad = sCad & Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "  (FIRMA)" & Space(nSep) & Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "  (FIRMA)" & Chr$(10)

sCad = sCad & sBON & Justifica("---- BENEFICIARIO LA TRANSACCIÓN ------", nMax) & sBOFF & Space(nSep) & sBON & Justifica("---- BENEFICIARIO LA TRANSACCIÓN ------", nMax) & sBOFF & Chr$(10)
sCad = sCad & Justifica("NOMBRE : " & sBenNombre, nMax) & Space(nSep) & Justifica("NOMBRE : " & sBenNombre, nMax) & Chr$(10)
sCad = sCad & Justifica("DOC ID : " & sBenDocId, nMax) & Space(nSep) & Justifica("DOC ID : " & sBenDocId, nMax) & Chr$(10)
sCad = sCad & Justifica("DIREC. : " & sBenDireccion, nMax) & Space(nSep) & Justifica("DIREC. : " & sBenDireccion, nMax) & Chr$(10)

If Len(sCuenta) = 18 Then
    sOpe = sOperacion & " (" & IIf(CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional, "SOLES", "DOLARES") & ")"
Else
    If nMoneda = 1 Then
        sOpe = sOperacion & " (SOLES)"
    ElseIf nMoneda = 2 Then
        sOpe = sOperacion & " (DOLARES)"
    End If
End If

If Len(sOpe) <= nMax Then
   sCad = sCad & sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF & Space(nSep) & sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF & Chr$(10)
Else
   sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES") & ")"
   sCad = sCad & sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF & Space(nSep) & sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF & Chr$(10)
End If

sCad = sCad & sBON & Justifica("MONTO", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "##,##0.00"), 14) & sBOFF & Space(nSep) & sBON & Justifica("MONTO", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "##,##0.00"), 14) & sBOFF & Chr$(10)
sCad = sCad & sBON & Justifica("TIPO CUENTA", nMax - 15) & " " & Right(Space(14) & sTipo, 14) & sBOFF & Space(nSep) & sBON & Justifica("TIPO CUENTA", nMax - 15) & " " & Right(Space(14) & sTipo, 14) & sBOFF & Chr$(10)
sCad = sCad & String(nMax, "=") & Space(nSep) & String(nMax, "=") & Chr$(10)
sCad = ImpreCarEsp(sCad)

sCad = sCad & Chr$(10)
If bCadena Then
    ImprimeBoletaLavadoDinero = sCad
Else
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sCad
    Close nFicSal
End If
Exit Function
ImpreErr:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function


Public Function ImprimeBoletaDepCtaBanco(ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal dFecha As Date, ByVal sBanco As String, ByVal sTipoCuenta As String, _
        ByVal sDocumento As String, ByVal nImporte As Double, ByVal sLpt As String, _
        ByVal nMoneda As Moneda, ByVal sUsuario As String, ByVal sOperacion As String, _
        Optional bCadena As Boolean = False) As String
    
Dim sImpre As String, sCad As String
Dim nFicSal As Integer
Dim sBON As String, sBOFF As String, sOpe As String
Dim nFila As Integer, nSep As Integer, nMax As Integer
Dim sTexto As String

ETIQ:
    On Error GoTo ImpreErr
    

sCad = sCad & Chr$(27) & Chr$(64)

sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita
    
nSep = 14
nMax = 39

sBON = Chr$(27) & Chr$(71)
sBOFF = Chr$(27) & Chr$(72)

sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & LinRec(sBON & Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & sBOFF, nSep)
sCad = sCad & LinRec(sBON & Justifica(Trim(sNomAge), nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & sBOFF, nSep)
sCad = sCad & Chr$(10)
sCad = sCad & LinRec(String(nMax, "-"), nSep)
sCad = sCad & LinRec(sBON & Justifica(sBON & "Entidad: " & sBanco & sBOFF, nMax + 2) & sBOFF, nSep)
sCad = sCad & LinRec(sBON & Justifica(sBON & "Cuenta : " & sTipoCuenta & sBOFF, nMax + 2) & sBOFF, nSep)
sCad = sCad & LinRec(sBON & Justifica(sBON & "Doc    : " & sDocumento & sBOFF, nMax + 2) & sBOFF, nSep)
sCad = sCad & LinRec(String(nMax, "-"), nSep)
sCad = sCad & Chr$(10)
sOpe = sOperacion & " (" & IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES") & ")"

If Len(sOpe) <= nMax Then
   sCad = sCad & LinRec(sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF, nSep)
Else
   sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES") & ")"
   sCad = sCad & LinRec(sBON & Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & sBOFF, nSep)
End If
sCad = sCad & Chr$(10)
sCad = sCad & LinRec(sBON & Justifica("MONTO", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "#,##0.00"), 14) & sBOFF, nSep)
sCad = sCad & LinRec(String(nMax, "="), nSep)
sCad = sCad & LinRec(Justifica(sUsuario, nMax), nSep)

sCad = ImpreCarEsp(sCad)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(10)

If bCadena Then
    ImprimeBoletaDepCtaBanco = sCad
Else
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sCad
    Close nFicSal
End If
Exit Function
ImpreErr:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function



