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
Dim I As Integer
Dim sCad As String

sFchLet = ArmaFecha(dApertura)
dVencimiento = DateAdd("d", nPlazo, dApertura)
sVctLet = ArmaFecha(dVencimiento)
sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", "US$ ")

sTaInNo = Format$((nTasa / 12), "##0.0000")
sTaInEf = Format$(((((((nTasa / 12) / 3000) + 1) ^ 30) - 1) * 100), "##0.0000")
sTaInEf = Left(sTaInEf, (Len(sTaInEf) - 2))
sTaInAn = Format(((((((nTasa / 12) / 3000) + 1) ^ 360) - 1) * 100), "##0.0000")

sPlazo = Trim(nPlazo) & " dia(s)"


sCad = sCad & Chr$(15)                           'Retorna al tipo de letra normal
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
sCad = sCad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
sCad = sCad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
sCad = sCad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna

For I = 1 To 6
    sCad = sCad & Chr$(10)     '   "Fila No. " & Trim(Str(i))
Next I

'       Impresión
sCad = sCad & Space(43) & UCase(sFchLet) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & ImpreCarEsp(sNomCli) & Chr$(10)      '  psNomCli
sCad = sCad & sDirCli & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "PLAZO FIJO " & Space(24) & sCuenta & Space(57) & FillNum(Trim(nExtracto), 5, " ") & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "DEPOSITO A PLAZO FIJO " & Space(9) & ":" & Space(3) & Trim(nPlazo) & Space(54) & Chr$(27) + Chr$(69) & sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2) & Chr$(27) + Chr$(70)
sCad = sCad & "FECHA DE VENCIMIENTO               "
sCad = sCad & "Y/O CANCELACION                :   " & UCase(sVctLet) & Chr$(10)
sCad = sCad & "TASA DE INTERES NOMINAL        :   " & Left(sTaInNo, (Len(sTaInNo) - 2)) & "%  MENSUAL" & Chr$(10)
sCad = sCad & "TASA DE INTERES EFECTIVA       :   " & sTaInEf & "%  MENSUAL " & Chr$(10)
sCad = sCad & "FRECUENCIA DE CAPITALIZACION   :   DIARIA" & Chr$(10)
sCad = sCad & "LOS INTERESES SE ABONARAN      :   " & Left(sRetInt, 20) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & Chr$(27) + Chr$(69)                           'Establece tipo de letra negrita
sCad = sCad & "SON :" & Chr$(10)
sCad = sCad & sLetras
sCad = sCad & Space(54) & sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2) & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & "TASA DE INTERES EFECTIVA ANUAL :   " & Left(sTaInAn, (Len(sTaInAn) - 2)) & "%"
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
        Optional sLpt As String = "", Optional bCadena As Boolean = False, Optional bPreguntaSaldo As Boolean = True) As String

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
sAux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

'scad = scad & Chr$(10);
sCad = sCad & lsNegritaOn 'Activa Negrita
sCad = sCad & Space(sIni) & "CMACT - AHORRO" & Space(19 + sSep + sAux) & "CMACT - AHORRO" & Chr$(10)

If Mid(psCodCta, 9, 1) = 1 Then
    sCad = sCad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
Else
    sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
End If

If psNomAgeRem = "" Then
    sCad = sCad & Chr$(10)
Else
    sCad = sCad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        sCad = sCad & lsNegritaOff 'Desactiva Negrita
    Else
        sCad = sCad & Space(sIni) & psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff & Chr$(10)  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    sCad = sCad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
    lnNumLinCmac = 1
End If

sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sAux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Chr$(10)

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
    sCad = sCad & Space(sIni) & ImpreCarEsp(psCliente) & Space(sAux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & Chr$(10)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
    
    sCad = sCad & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(sAux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & Chr$(10)
    sCad = sCad & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(sAux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & Chr$(10)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    sCad = sCad & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + sAux) & "Cuenta:" & psCodCta & Chr$(10)
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
    sCad = sCad & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(sAux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & Chr$(10)
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
    lnTope = 4 - lnCliAux
Else
    lnTope = 3 - lnCliAux
End If

sCad = sCad & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
If psCodUsuRem = "" Then
    sCad = sCad & Space(sIni) & ImpreCarEsp(sCodUser) & Space(29 + sSep + sAux) & ImpreCarEsp(sCodUser) & Chr$(10)
Else
    sCad = sCad & Space(sIni) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Chr$(10)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
sCad = sCad & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & ImpreCarEsp(lsMensaje) & lsNegritaOff & Chr$(10)

lnNumLinCmac = lnNumLinCmac + 1
 
'scad = scad & Chr$(12)
For sAux = 1 To (lnTope - lnNumLinCmac)
    sCad = sCad & Chr$(10)
Next sAux

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
ERROR:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function

