VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sCodUser As String
Dim CuentasPersonaPrd() As String
Dim CuentasPersonaCre() As String
Dim CuentasPersonaJud() As String

Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "") As String

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
Open sLpt For Output As nFicSal

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

lsNomAge = sNomAge

'Print #nFicSal, Chr$(10);
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"

If Mid(psCodCta, 9, 1) = 1 Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
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
    If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
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
    Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + sAux); ImpreCarEsp(sCodUser)
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
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

Public Sub ImprimeBoletaInteres(ByVal psTit As String, ByVal psText As String, _
        ByVal psCodOpe As String, ByVal pnMonto As String, ByVal psCliente As String, _
        ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, ByVal NomDoc As String, ByVal pnNumExt As Long, _
        ByVal psTextInt As String, ByVal psInter As String, Optional pbSaldoCnt As Boolean = False, _
        Optional psSaldoCnt As String = "XXX", Optional pbPSaldos As Boolean = True, _
        Optional sNomAge As String = "", Optional dFecSis As Date, Optional sCodUser As String = "", _
        Optional sLpt As String)

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax  As Integer
Dim sAux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lsMonInt As String
Dim lsMonTotal As String
Dim LnMonTotal As Currency

Dim lsCliAux1 As String
Dim lsCliAux2 As String
Dim lnCliAux As Long
Dim lsSaldCnt As String

Dim lsMensaje As String * 39

ETIQ:
On Error GoTo ERROR

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open sLpt For Output As nFicSal

Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
sAux = 5


sFecha = Format$(dFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "###,##0.00")
sSDisp = Format$(pnSaldo, "###,##0.00")

lsMonInt = Format$(psInter, "###,##0.00")
LnMonTotal = Val(psInter) + Val(pnMonto)
lsMonTotal = Format(LnMonTotal, "###,##0.00")

Dim lsCadArg As String
Dim lsInteres As String

'Print #nFicSal, Chr$(10);
Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni); ""
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"

If Mid(psCodCta, 6, 1) = 1 Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(sMax + sAux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(sMax + sAux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt;
End If

Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

Print #nFicSal, ""
Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sMax - Len(psCliente) < 2 Then psCliente = Mid(psCliente, 1, sMax - 2)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(sAux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax - 2)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax - 2)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    lnCliAux = 1
End If

Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta

psTit = Trim(psTit)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"


Print #nFicSal, Tab(sIni); ImpreCarEsp(psText); Space(sMax + 6 - Len(psText) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(psText); Space(sMax + 6 - Len(psText) - Len(sMonto)); sMonto
Print #nFicSal, Tab(sIni); ImpreCarEsp(psTextInt); Space(sMax + 6 - Len(psTextInt) - Len(lsMonInt)); lsMonInt; Space(-1 + sSep); ImpreCarEsp(psTextInt); Space(sMax + 6 - Len(psTextInt) - Len(lsMonInt)); lsMonInt
Print #nFicSal, Tab(sIni); "Total"; Space(sMax + 6 - Len("Total") - Len(lsMonTotal)); lsMonTotal; Space(-1 + sSep); "Total"; Space(sMax + 6 - Len("Total") - Len(lsMonTotal)); lsMonTotal
Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

Print #nFicSal, ""


If pbPSaldos Then
    If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
        If pbSaldoCnt Then
            lsSaldCnt = Format(psSaldoCnt, "#,##0.00")
            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldCnt)); lsSaldCnt; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldCnt)); lsSaldCnt
            lnCliAux = lnCliAux + 1
        End If
    Else
        Print #nFicSal, ""
        lsInteres = "No Valido"
    End If
Else
    Print #nFicSal, ""
End If
lsInteres = pnInteresA

If lsInteres <> "No Valido" Then
    lsInteres = Format(lsInteres, "###,##0.00")
    Print #nFicSal, Tab(sIni); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres; Space(-1 + sSep); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres
Else
    Print #nFicSal, ""
End If


If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + sAux); ImpreCarEsp(sCodUser)

Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing

Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
lnCliAux = lnCliAux + 1

'Print #nFicSal, Chr$(12)
For sAux = 1 To 4 - lnCliAux
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

Private Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bBoletaInt As Boolean = False, _
            Optional nIntCanc As Double = 0, Optional nDias As Long = 0, Optional bImpSaldos As Boolean = True) As String

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
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos
            Case TpoDocNotaAbono, TpoDocNotaCargo
                If bBoletaInt Then
                    ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, sNroDoc, nSaldoDisp, Trim(nIntMes), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntCanc), , , bImpSaldos
                Else
                    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos
                End If
            Case TpoDocOrdenPago
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos
        End Select
    Else
        If bBoletaInt Then
            ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, Trim(nIntMes), "", nExtracto, "Interes Retirado", Trim(nIntCanc), , , bImpSaldos
        Else
            If nDias > 0 Then
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", 1, nSaldoCnt, bImpSaldos, , Trim(nDias)
            Else
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", 1, nSaldoCnt, bImpSaldos
            End If
        End If
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

'Public Function ValidaMontoApertura(ByVal nProducto As Producto, ByVal nPersoneria As PersPersoneria, _
'        ByVal nMonto As Double, ByVal nMoneda As Moneda, Optional bOrdPag As Boolean = False) As String
'Dim nValor As Double
'Dim clsGen  As nCapDefinicion
'Dim sDesc As String
'Set clsGen = New nCapDefinicion
'nValor = clsGen.GetMontoMinimoAperturaPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
'Set clsGen = Nothing
'If nValor > nMonto Then
'    ValidaMontoApertura = "Monto de Apertura debe ser > ó = a " & Format$(nValor, "#,##0.00")
'Else
'    ValidaMontoApertura = ""
'End If
'End Function

'Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsPers As Recordset, _
'        ByVal sAgencia As String, ByVal nOperacion As CaptacOperacion, ByVal nTasa As Double, ByVal nSaldo As Double, _
'        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, _
'        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
'        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
'        Optional sNroDoc As String, Optional sCodIF As String, _
'        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
'        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
'        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
'        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
'        Optional sPersLavDinero As String = "", Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional) As String
'
'Dim clsMant As DCapMovimientos
'Dim nErrNum As Integer
'Dim nErrDesc As String, sCuenta As String
'Dim bTrans As Boolean
'Dim nMovNro As Long
'Dim lnMontoCV As Currency
'
'bTrans = False
'Set clsMant = New DCapMovimientos
'On Error GoTo ErrGraba
'clsMant.dbCmact.BeginTrans
'bTrans = True
'sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
'    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
'clsMant.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsMant.GetnMovNro(sMovNro)
'Select Case nOperacion
'    Case gAhoApeChq, gPFApeChq, gCTSApeChq
'        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
'        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
'        clsMant.AgregaCuentaDocumento sCuenta, TpoDocCheque, sNroDoc, sCodIF, sMovNro, nMovNro, nSaldo
'    Case gAhoApeTransf, gPFApeTransf, gCTSApeTransf
'        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
'        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
'        'clsMant.AgregaCuentaDocumento sCuenta, TpoDocNotaAbono , sNroDoc, sCodIF, sMovNro, nMovNro, nSaldo
'
'        If pnMovNroTransfer <> -1 Then
'            clsMant.AgregaMovRef nMovNro, pnMovNroTransfer
'            If pnMonedaTrans <> nMoneda Then
'                GetTipCambio dFecha
'
'                If nMoneda = gMonedaNacional Then
'                    clsMant.ActualizaMovPendientesRend pnMovNroTransfer, Round(nSaldo / gnTipCambioC, 2)
'                    lnMontoCV = Round((nSaldo / gnTipCambioC) * gnTipCambio, 2)
'                    clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
'                    clsMant.AgregaMovTipoCambio nMovNro, gnTipCambioC
'                    If gnTipCambioC < gnTipCambio Then
'                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nSaldo)
'                    Else
'                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nSaldo)
'                    End If
'                Else
'                    clsMant.ActualizaMovPendientesRend pnMovNroTransfer, Round(nSaldo * gnTipCambioV, 2)
'                    lnMontoCV = Round(nSaldo * gnTipCambio, 2)
'                    clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
'                    clsMant.AgregaMovTipoCambio nMovNro, gnTipCambioV
'                    If gnTipCambioV > gnTipCambio Then
'                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nSaldo * gnTipCambioV, 2))
'                    Else
'                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nSaldo * gnTipCambioV, 2))
'                    End If
'                End If
'            Else
'                clsMant.ActualizaMovPendientesRend pnMovNroTransfer, nSaldo
'            End If
'
'        End If
'        'Aqui debe registrarse la referencia con la transferencia de caja general
'    Case Else
'        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
'        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
'End Select
'If nOperacion = gCTSApeEfec Or nOperacion = gCTSApeChq Or nOperacion = gCTSApeTransf Then
'    clsMant.AgregaMovParamCTS nMovNro, nPorcRetCTS
'End If
'If sPersLavDinero <> "" Then
'    clsMant.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
'Do While Not rsPers.EOF
'    clsMant.AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4)))
'    rsPers.MoveNext
'Loop
'clsMant.dbCmact.CommitTrans
'CapAperturaCuenta = sCuenta
'Exit Function
'ErrGraba:
'    If bTrans Then clsMant.dbCmact.RollbackTrans
'    Set clsMant = Nothing
'    CapAperturaCuenta = ""
'    Err.Raise Err.Number, "CapAperturaCuenta", Err.Description
'End Function

'Public Sub CapAperturaCuentaLote(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsCuenta As Recordset, _
'    ByVal nOperacion As CaptacOperacion, ByVal sGlosa As String, ByVal sMovNro As String, ByVal nTipoTasa As CaptacTipoTasa, _
'    Optional bCheque As Boolean = False, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
'    Optional dFechaValorizacion As Date, Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
'    Optional sNomCMAC As String = "", Optional sNomAge As String = "", Optional sLpt As String = "")
'
'Dim clsMov As DCapMovimientos
'Dim clsMant As NCapMantenimiento
'Dim nErrNum As Integer
'Dim nErrDesc As String, sCuenta As String, sCtaCodAbono As String
'Dim bTrans As Boolean, bOrdPag As Boolean, bCtaAboInt As Boolean
'Dim nMovNro As Long, nPlazo As Long, nFirmas As Long
'Dim nSaldo As Double, nTasa As Double
'Dim nFormaRetiro As CaptacPFFormaRetiro
'Dim nPersoneria As PersPersoneria
'Dim nTipoCuenta As ProductoCuentaTipo
'Dim sImpresion As String, sCodUser As String, sAgencia As String
'Dim dFecSis As Date
'Dim sCad As String, sNomTit As String
'Dim clsImp As NCapImpBoleta
'Dim rsProdPers As Recordset
'Dim lbImpRegFirma As Byte
'
'Dim oGen As DGeneral
'Set oGen = New DGeneral
'
'lbImpRegFirma = CInt(oGen.LeeConstSistema(100))
'Set oGen = Nothing
'
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'sAgencia = Mid(sMovNro, 18, 2)
'
'bTrans = False
'Set clsMov = New DCapMovimientos
'On Error GoTo ErrGraba
'
'
'clsMov.dbCmact.BeginTrans 'Inicializa la transacción
'bTrans = True
'
''Agrega el movimiento
'clsMov.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsMov.GetnMovNro(sMovNro)
'If nOperacion = gCTSApeLoteEfec Or nOperacion = gCTSApeLoteChq Then
'    clsMov.AgregaMovParamCTS nMovNro, nPorcRetCTS
'End If
'clsMov.dbCmact.CommitTrans 'Inicializa la transacción
'bTrans = False
'
'
''Variables globales para las aperturas en lote
'bCtaAboInt = False
'nTipoCuenta = gPrdCtaTpoIndiv
'sCtaCodAbono = ""
'nFirmas = 1
'
'Set clsMant = New NCapMantenimiento
'If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
'    clsMov.AgregaMovDoc nMovNro, TpoDocCheque, sNroDoc, dFecSis
'End If
'Do While Not rsCuenta.EOF
'    clsMov.dbCmact.BeginTrans
'    bTrans = True
'    nSaldo = CDbl(rsCuenta("Monto"))
'    nPersoneria = CInt(rsCuenta("cPersoneria"))
'    nTasa = CDbl(rsCuenta("Tasa"))
'    nPlazo = 0
'    bOrdPag = False
'    If nProducto = gCapPlazoFijo Then
'        nPlazo = rsCuenta("Plazo")
'        nFormaRetiro = CLng(Right(rsCuenta("Forma Retiro"), 4))
'    ElseIf nProducto = gCapAhorros Then
'        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
'    End If
'
'    sCuenta = clsMov.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecSis, nFirmas, nPersoneria, _
'        nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
'    rsCuenta("Cuenta") = sCuenta
'    rsCuenta.Update
'    Select Case nOperacion
'        Case gAhoApeLoteChq, gPFApeLoteChq, gCTSApeLoteChq
'            clsMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
'            clsMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
'            clsMov.AgregaDocRecCapta sCuenta, TpoDocCheque, sNroDoc, sCodIF, nMovNro, nSaldo
'        Case Else
'            clsMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
'            clsMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
'    End Select
'
'    clsMov.AgregaNuevoProdPers sCuenta, rsCuenta("Código"), gCapRelPersTitular
'    clsMov.dbCmact.CommitTrans
'    bTrans = False
'    rsCuenta.MoveNext
'Loop
'
'Set clsMov = Nothing
'sCad = ""
'
''Imprime registro de firmas
'rsCuenta.MoveFirst
'Do While Not rsCuenta.EOF
'    bOrdPag = False
'    If nProducto = gCapAhorros Then
'        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
'    End If
'    If sCad <> "" Then sCad = sCad & Chr$(12)
'    'Creamos el recordset para el registro de firmas
'    Set rsProdPers = New Recordset
'    rsProdPers.Fields.Append "Codigo", adVarChar, 13
'    rsProdPers.Fields.Append "Relacion", adVarChar, 30
'    rsProdPers.Open
'    rsProdPers.AddNew
'    rsProdPers("Codigo") = rsCuenta("Código")
'    rsProdPers("Relacion") = "TITULAR"
'    rsProdPers.Update
'    If lbImpRegFirma = 1 Then
'        sCad = sCad & clsMant.GeneraRegistroFirmas(rsCuenta("Cuenta"), "INDIVIDUAL", dFecSis, bOrdPag, rsProdPers, sNomAge, dFecSis)
'    End If
'    Set rsProdPers = Nothing
'    rsCuenta.MoveNext
'Loop
'
'Set clsImp = New NCapImpBoleta
'
''Imprime Certificado de plazos fijos
'If nProducto = gCapPlazoFijo Then
'    Dim sDirCli As String
'    Dim sFormaRetiro As String
'    Dim sLetras As String
'    Dim rsPers As Recordset
'    Dim clsBol As NCapImpBoleta
'    rsCuenta.MoveFirst
'    Do While Not rsCuenta.EOF
'        nSaldo = CDbl(rsCuenta("Monto"))
'        nTasa = CDbl(rsCuenta("Tasa"))
'        nPlazo = rsCuenta("Plazo")
'        nFormaRetiro = CLng(Right(rsCuenta("Forma Retiro"), 4))
'        sNomTit = clsMant.GetNombreTitulares(rsCuenta("Cuenta"))
'        Set rsPers = clsMant.GetDatosPersona(rsCuenta("Código"))
'        sDirCli = Trim(rsPers("Direccion"))
'        sLetras = ConversNL(nMoneda, nSaldo)
'        sFormaRetiro = Trim(Left(rsCuenta("Forma Retiro"), 25))
'        If sCad <> "" Then sCad = sCad & Chr$(12)
'        sCad = sCad & clsImp.ImprimeCertificadoPlazoFijo(dFecSis, sNomTit, sDirCli, rsCuenta("Cuenta"), "1", nPlazo, nSaldo, _
'                    nTasa, sFormaRetiro, sLetras, , True)
'        rsCuenta.MoveNext
'    Loop
'End If
'
''Imprime las boletas
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'sTipDep = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES")
'If bCheque Then
'    sModDep = "Depósito Cheque"
'Else
'    sModDep = "Depósito Efectivo"
'End If
'Select Case nProducto
'    Case gCapAhorros
'        sTipApe = "APERTURA AHORROS"
'    Case gCapPlazoFijo
'        sTipApe = "APERTURA PLAZO FIJO"
'    Case gCapCTS
'        sTipApe = "APERTURA CTS"
'End Select
'rsCuenta.MoveFirst
'Do While Not rsCuenta.EOF
'    bOrdPag = False
'    If nProducto = gCapAhorros Then
'        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
'    End If
'    If bOrdPag Then sTipApe = "APERTURA AHORROS CON OP"
'    sNomTit = clsMant.GetNombreTitulares(rsCuenta("Cuenta"))
'    If sCad <> "" Then sCad = sCad & Chr$(12)
'    If bCheque Then
'        sCad = sCad & clsImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, Trim(nOperacion), Trim(nSaldo), sNomTit, rsCuenta("Cuenta"), Format$(dFechaValorizacion, "dd/mm/yyyy"), nSaldo, 0, "Fecha Valor", 1, nSaldo, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True)
'    Else
'        sCad = sCad & clsImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), Trim(sCodOpe), Trim(nSaldo), sNomTit, rsCuenta("Cuenta"), "", nSaldo, 0, "", 1, nSaldo, , , , , , , , , , dFecSis, Trim(sNomAge), sCodUser, sLpt, True)
'    End If
'    rsCuenta.MoveNext
'Loop
'Set clsImp = Nothing
'Set clsMant = Nothing
'
'Dim clsPrev As Previo.clsPrevio
'Set clsPrev = New Previo.clsPrevio
'clsPrev.Show sCad, "APERTURA EN LOTE", False
'Exit Sub
'ErrGraba:
'    If bTrans Then clsMov.dbCmact.RollbackTrans
'    Set clsMov = Nothing
'    MsgBox Err.Description, vbExclamation, "Error"
'End Sub

Public Function ValidaCuentaOperacion(ByVal sCuenta As String, Optional bAbono As Boolean = False) As String
Dim clsCap As DCapMovimientos
Dim nEstado As CaptacEstado
Dim rsBlq As Recordset
Dim sCad As String
Set clsCap = New DCapMovimientos
nEstado = clsCap.GetEstadoCuenta(sCuenta)
If nEstado < 0 Then
    ValidaCuentaOperacion = "Cuenta No Existe"
Else
    Select Case nEstado
        Case gCapEstAnulada, gCapEstCancelada
            ValidaCuentaOperacion = "Cuenta CANCELADA O ANULADA"
        Case gCapEstBloqTotal, gCapEstBloqRetiro
            Set rsBlq = clsCap.GetCuentaBloqueos(sCuenta, gCapTpoBlqRetiro)
            sCad = ""
            If Not bAbono Then
                If Not (rsBlq.EOF And rsBlq.BOF) Then
                    sCad = "BLOQUEO PARA RETIRO: " & Chr$(13)
                    Do While Not rsBlq.EOF
                        sCad = sCad & "- " & Trim(rsBlq("Motivo")) & Chr$(13)
                        rsBlq.MoveNext
                    Loop
                End If
            End If
            Set rsBlq = clsCap.GetCuentaBloqueos(sCuenta, gCapTpoBlqTotal)
            If Not (rsBlq.EOF And rsBlq.BOF) Then
                sCad = sCad & "BLOQUEO TOTAL: " & Chr$(13)
                Do While Not rsBlq.EOF
                    sCad = sCad & "- " & Trim(rsBlq("Motivo")) & Chr$(13)
                    rsBlq.MoveNext
                Loop
            End If
            If sCad <> "" Then
                ValidaCuentaOperacion = sCad
            End If
        Case Else
            ValidaCuentaOperacion = ""
    End Select
End If
End Function

Public Function ValidaRetiroIntPF(ByVal sCuenta As String, ByVal dFecSis As Date) As String
Dim clsCap As DCapMantenimiento
Dim sCad As String
Dim dAuxiliar As Date, dRenovacion As Date
Dim nPlazo As Long
Dim rsCta As Recordset
Set clsCap = New DCapMantenimiento
Set rsCta = clsCap.GetDatosCuentaPF(sCuenta)
Set clsCap = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    dAuxiliar = rsCta("dAuxiliar")
    dRenovacion = rsCta("dRenovacion")
    nPlazo = rsCta("nPlazo")
    If DateDiff("d", dRenovacion, dFecSis) < 30 Then
        sCad = "El número de dias transcurridos desde la APERTURA es MENOR A 30 DIAS"
    Else
        If DateDiff("d", dAuxiliar, dFecSis) < 30 Then
            sCad = "El número de dias transcurridos desde la última " & Chr$(13) _
                    & "FECHA PACTADA DE RETIRO es MENOR A 30 DIAS"
        Else
            sCad = ""
        End If
    End If
Else
    sCad = "Cuenta NO Existe"
End If
rsCta.Close
Set rsCta = Nothing
ValidaRetiroIntPF = sCad
End Function

Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim nProducto As Producto
Dim clsMant As NCapMantenimiento
Dim clsPar As nCapDefinicion
Dim rsCta As Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Dim bOrdPag As Boolean
Dim nPersoneria As PersPersoneria
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
nProducto = CLng(Mid(sCuenta, 6, 3))
bOrdPag = False
If nProducto = gCapAhorros Then
    nSaldo = rsCta("nSaldoDisp")
    bOrdPag = rsCta("bOrdPag")
ElseIf nProducto = gCapCTS Then
    nSaldo = rsCta("nSaldRetiro")
End If
nPersoneria = rsCta("nPersoneria")
rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))

'If CLng(Mid(sCuenta, 6, 3)) = gCapAhorros Then
'    Set clsPar = New nCapDefinicion
'    nMontoMinimo = clsPar.GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
'    Set clsPar = Nothing
'Else
    nMontoMinimo = 0
'End If
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function

Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
        Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
        Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
        Optional sCodCMAC As String = "", Optional sNomCMAC As String = "", Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
        Optional bImprimeBoleta As Boolean = True, Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional pCon As ADODB.Connection = Nothing, Optional pbImprimeBoleta As Boolean = True) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean, bInamovilizada As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String, sCodUser As String
Dim nMovNro As Long
Dim dFecSis As Date
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Dim lnMontoCV As Currency


dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
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
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
bInamovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
If Not bTransaccion Then clsCap.dbCmact.BeginTrans
If Not pCon Is Nothing Then clsCap.SetConexion pCon

bTrans = True
If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)

If sNroDoc = "" Then ' Si la operacion es en efectivo
    clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto
    clsCap.ActualizadUltContactoAho sCuenta, dFecSis
Else
    If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
        nSaldoDisp = nSaldoDisp + 0
        nIntGanado = 0
        clsCap.ActualizadUltContactoAho sCuenta, dFecSis
    ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
        clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
    End If
End If
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

If sNroDoc = "" Then ' Si la operacion es en efectivo
    clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    If nOperacion = gAhoDepEfec Then
        sMsgOpe = "Depósito Efectivo"
    ElseIf nOperacion = gAhoDepTransf Then
        sMsgOpe = "Depósito Transferencia"
    ElseIf nOperacion = gAhoDepDevFondoGar Then
        sMsgOpe = "Depósito Dev Fondo Gar"
    End If
Else
    If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto
        sMsgOpe = "Depósito Cheque"
    ElseIf nTipoDoc = TpoDocNotaAbono Then  'Si el abono es con NA
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        'clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto
        sMsgOpe = "Depósito Nota Abono"
    End If
End If

If bInamovilizada Then
    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInmovilizadaAct sCuenta
    End If
End If

If bInactiva And Not bInamovilizada Then
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    clsCap.ActualizaInactivaAct sCuenta
End If

If pnMovNroTransfer <> -1 Then
    clsCap.AgregaMovRef nMovNro, pnMovNroTransfer
    If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis, Not IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
        
        If Mid(sCuenta, 9, 1) = gMonedaNacional Then
            clsCap.ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto / gnTipCambioC, 2)
            lnMontoCV = Round((nMonto / gnTipCambioC) * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioC
            If gnTipCambioC < gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
            End If
        Else
            clsCap.ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto * gnTipCambioV, 2)
            lnMontoCV = Round(nMonto * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioV
            If gnTipCambioV > gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
            End If
        End If
    Else
        clsCap.ActualizaMovPendientesRend pnMovNroTransfer, nMonto
    End If
End If

If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
If Not bTransaccion Then clsCap.dbCmact.CommitTrans
bTrans = False
CapAbonoCuentaAho = nSaldoCnt

If Not bImprimeBoleta Then Exit Function
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "DEPOSITO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

If pbImprimeBoleta Then
    Do
        If sNroDoc <> "" Then
            Select Case nTipoDoc
                Case TpoDocCheque
                    If sCodCMAC <> "" Then
                        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
                    Else
                        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
                    End If
                Case TpoDocNotaAbono
                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
            End Select
        Else
            If sCodCMAC <> "" Then
                sTipApe = "DEPOSITO CMAC AHORROS"
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
            Else
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
            End If
        End If
    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
End If
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function

Public Function CapAbonoCuentaCTS(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nPorcDisp As Double = 0, Optional nTipoDoc As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCMAC As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional pnTipoCambio As Currency = 0, _
            Optional nIntagibleCTS As Currency = 0) As Double 'MAVM 20120504

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nSaldoRetiro As Double
Dim nTotalSaldos As Double, nTotalSueldos As Double, nExcedente As Double 'BRGO 20110428
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bTrans As Boolean, bInactiva As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim lnMontoCV As Currency
Dim nDU01 As Double 'APRI20200415 POR COVID-19
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nSaldoRetiro = rsCta("nSaldRetiro")
nExtracto = rsCta("nTransacc") + 1
nDU01 = rsCta("nDU01") 'APRI20200415 POR COVID-19
rsCta.Close
Set rsCta = Nothing
If nPorcDisp = 0 Then
    Dim clsDef As nCapDefinicion
    Set clsDef = New nCapDefinicion
    nPorcDisp = clsDef.GetCapParametro(gPorRetCTS)
    Set clsDef = Nothing
End If
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
'Calcula Intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntSaldo = GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nSaldoCnt = nSaldoCnt + nMonto
dUltMov = DateAdd("d", -1, dFecSis)

If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)

'If sNroDoc = "" Then 'Si es abono en efectivo
'    nSaldoDisp = nSaldoDisp + nMonto
'    nSaldoRetiro = nSaldoRetiro + (nMonto * nPorcDisp / 100)
'Else
'    If nTipoDoc = TpoDocCheque Then 'Si es con cheque
'        nSaldoDisp = nSaldoDisp + 0
'    Else
'        nSaldoRetiro = nSaldoRetiro + (nMonto * nPorcDisp / 100)
'    End If
'End If
'*** BRGO 20110428 **** Ley de disponibilidad e intangibilidad de la CTS
Set rsCta = clsCap.GetCapRHSaldosYSueldosCTS(sCuenta, dFecSis, pnTipoCambio)
nTotalSaldos = rsCta!TotalSaldos
'MAVM 20120504 ***
'nTotalSueldos = rsCta!totalsueldos
nTotalSueldos = nIntagibleCTS
'***
nExcedente = (nTotalSaldos + nMonto) - nTotalSueldos
rsCta.Close
If nExcedente <= 0 Then
    nExcedente = 0
End If

If sNroDoc = "" Then 'Si es abono en efectivo
    nSaldoDisp = nSaldoDisp + nMonto
    nSaldoRetiro = nExcedente * nPorcDisp / 100
Else
    If nTipoDoc = TpoDocCheque Then 'Si es con cheque
        nSaldoDisp = nSaldoDisp + 0
    Else
        nSaldoRetiro = nExcedente * nPorcDisp / 100
    End If
End If
'*** End BRGO *************************************************************

clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvision, nIntGanado 'COMENTADO POR APRI20180926 MEJORA CORREO-MARI 21/09/2018
clsCap.AgregaMovParamCTS nMovNro, nPorcDisp
If sNroDoc = "" Then 'Si es abono en efectivo
    'clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo
    clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
    clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    If nOperacion = gCTSDepEfec Then
        sMsgOpe = "Depósito Efectivo"
    ElseIf nOperacion = gCTSDepTransf Then
        sMsgOpe = "Depósito Transferencia"
    End If
Else
    If nTipoDoc = TpoDocCheque Then 'Si es con cheque
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto
        sMsgOpe = "Depósito Cheque"
    ElseIf nTipoDoc = TpoDocNotaAbono Then  'Si es con Nota Abono
        'clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo
        clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        'clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto
        sMsgOpe = "Depósito Nota Abono"
    End If
End If


If pnMovNroTransfer <> -1 Then
    clsCap.AgregaMovRef nMovNro, pnMovNroTransfer
    If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis, Not IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
        
        If Mid(sCuenta, 9, 1) = gMonedaNacional Then
            clsCap.ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto / gnTipCambioC, 2)
            lnMontoCV = Round((nMonto / gnTipCambioC) * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioC
            If gnTipCambioC < gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
            End If
        Else
            clsCap.ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto * gnTipCambioV, 2)
            lnMontoCV = Round(nMonto * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioV
            If gnTipCambioV > gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
            End If
        End If
    Else
        clsCap.ActualizaMovPendientesRend pnMovNroTransfer, nMonto
    End If
End If




If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
clsCap.dbCmact.CommitTrans
CapAbonoCuentaCTS = nSaldoCnt
bTrans = False

Dim sTipDep As String, sCodOpe As String
Dim sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sTipApe = "DEPOSITO CTS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
'comentado por gitu 08-05-2009 para que no pida la reimpresion de la boleta
'Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocCheque
                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
        End Select
    Else
        If sCodCMAC <> "" Then
            sTipApe = "DEPOSITO CMAC CTS"
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
        End If
    End If
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Function

Public Function CapCargoAhoOPCertifcada(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal sNroDoc As String, Optional sCodIF As String = "", _
            Optional sCodCMAC As String = "", Optional sNomCMAC As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset, rsPers As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sNombreIF As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
If sCodIF <> "" Then
    Set rsPers = clsMant.GetDatosPersona(sCodIF)
    sNombreIF = rsPers("Nombre")
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nExtracto = rsCta("nTransacc")
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
clsCap.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
If nOperacion = gAhoRetOPCert Then
    sMsgOpe = "Retiro OP Cert."
ElseIf nOperacion = gAhoRetOPCertCanje Then
    sMsgOpe = "Retiro OP Cert.Canje"
End If
sGlosa = sGlosa & ".Cargo Cuenta = " & sCuenta
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
clsCap.dbCmact.CommitTrans
CapCargoAhoOPCertifcada = nSaldoCnt

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
bReImp = False
Do
    If sCodCMAC <> "" Then
        sTipApe = "RETIRO AHORROS CMAC"
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoAhoOPCertifcada = 0
End Function

Public Function CapCargoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCMAC As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
            Optional psCtaBanco As String = "", Optional pnMoneda As Moneda = gMonedaNacional, Optional pCon As ADODB.Connection, Optional pbImprimeBoleta As Boolean = True) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean, bInmovilizada As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas
Dim lnMontoCV As Currency

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
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
bInmovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
If Not bTransaccion Then clsCap.dbCmact.BeginTrans
If Not pCon Is Nothing Then clsCap.SetConexion pCon

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)

clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto
clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            clsCap.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        End If
    End If
Else
    If nOperacion = gAhoDctoEmiOP Then
        sMsgOpe = "Retiro Emisión OP"
    ElseIf nOperacion = gAhoRetTransf Then
        sMsgOpe = "Retiro Transferencia"
        
    ElseIf nOperacion = gAhoRetEfec Then
        sMsgOpe = "Retiro Efectivo"
    ElseIf nOperacion = gAhoRetRetencionJudicial Then
        sMsgOpe = "Retiro Retención Jud."
    ElseIf nOperacion = gAhoRetDuplicadoTarj Then
        sMsgOpe = "Retiro Dup.Tarj."
    End If
End If
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.ActualizadUltContactoAho sCuenta, dFecSis

If bInmovilizada Then
    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInmovilizadaAct sCuenta
    End If
End If

If bInactiva And Not bInmovilizada Then
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    clsCap.ActualizaInactivaAct sCuenta
End If

clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
If nOperacion = gAhoRetTransf Then
    clsCap.AgregaMovTranferBanco nMovNro, sCodIF, psCtaBanco, pnMoneda

    If pnMoneda <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis, Not IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
        
        If Mid(sCuenta, 9, 1) = gMonedaNacional Then
            lnMontoCV = Round((nMonto / gnTipCambioV) * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioV
            If gnTipCambioV > gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
            End If
        Else
            lnMontoCV = Round(nMonto * gnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, gnTipCambioC
            If gnTipCambioC < gnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * gnTipCambioC, 2))
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * gnTipCambioC, 2))
            End If
        End If
    End If
        
End If

clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
If Not bTransaccion Then clsCap.dbCmact.CommitTrans
CapCargoCuentaAho = nSaldoCnt
bTrans = False

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing

If pbImprimeBoleta Then
    Do
        If sNroDoc <> "" Then
            Select Case nTipoDoc
                Case TpoDocOrdenPago
                    If sCodCMAC <> "" Then
                        sTipApe = "RETIRO AHORROS CMAC"
                        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
                    Else
                        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
                    End If
            End Select
        Else
            If sCodCMAC <> "" Then
                sTipApe = "RETIRO AHORROS CMAC"
                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
            Else
                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
            End If
        End If
    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
End If

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoCuentaAho = 0
End Function

Public Function CapCargoCuentaCTS(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCMAC As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldRetiro As Double, nIntRetiro As Double, nIntGanadoRet As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el cálculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nSaldRetiro = rsCta("nSaldRetiro")
nIntRetiro = rsCta("nIntSaldo")
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    CapCargoCuentaCTS = -1
    Exit Function
End If
sGlosa = sGlosa & "Cargo Cuenta = " & sCuenta

'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanadoRet = GetInteres(nSaldRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto
nSaldoDisp = nSaldoDisp - nMonto
clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
clsCap.ActualizaCargoCTS sCuenta, nMonto, nIntGanadoRet
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
    sMsgOpe = "Retiro Efectivo"
End If
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
clsCap.dbCmact.CommitTrans
CapCargoCuentaCTS = nSaldoCnt
bTrans = False
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim clsNeg As NCapMantenimiento
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sTipApe = "RETIRO CTS"
Set clsNeg = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsNeg.GetNombreTitulares(sCuenta))
Set clsNeg = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp - nMonto, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        If sCodCMAC <> "" Then
            sTipApe = "RETIRO CMAC CTS"
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp - nMonto, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp - nMonto, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoCuentaCTS = 0
End Function

'Public Function GetSaldoCancelacion(ByVal sCuenta As String, ByVal dFecSis As Date, ByVal sCodAge As String, _
'                Optional ByRef nTasaCancelacion As Double = 0) As Double
'Dim nProd As Producto
'Dim rsProd As Recordset
'Dim clsMant As DCapMantenimiento
'Dim nSaldoDisp As Double, nSaldoCnt As Double
'Dim nIntAcum As Double, nIntGanado As Double
'Dim nTasa As Double
'Dim dUltMov As Date
'Dim nDiasTranscurridos As Long
'nProd = CLng(Mid(sCuenta, 6, 3))
'Set clsMant = New DCapMantenimiento
'Select Case nProd
'    Case gCapAhorros, gCapCTS
'        If nProd = gCapAhorros Then
'            Set rsProd = clsMant.GetDatosCuentaAho(sCuenta)
'        Else
'            Set rsProd = clsMant.GetDatosCuentaCTS(sCuenta)
'        End If
'        Set clsMant = Nothing
'        nSaldoDisp = rsProd("nSaldoDisp")
'        nSaldoCnt = rsProd("nSaldo")
'        dUltMov = rsProd("dUltCierre")
'        nTasa = rsProd("nTasaInteres")
'        nIntAcum = rsProd("nIntAcum")
'        nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'        nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'        GetSaldoCancelacion = nSaldoDisp + nIntAcum + nIntGanado
'        rsProd.Close
'        Set rsProd = Nothing
'    Case gCapPlazoFijo
'        Dim nTasaDias As Double, nIntPagado As Double
'        Dim nDias As Long, nPlazo As Long, nPlazoCumplido As Long
'        Dim dVencimiento As Date, dApertura As Date
'        Dim nInt01 As Double, nInt02 As Double, nInt03 As Double
'        Dim clsCap As DCapMovimientos
'        Dim clsDef As nCapDefinicion
'        Set rsProd = clsMant.GetDatosCuentaPF(sCuenta)
'        Set clsMant = Nothing
'        nSaldoDisp = rsProd("nSaldoDisp")
'        nSaldoCnt = rsProd("nSaldo")
'        dApertura = rsProd("dRenovacion")
'        dUltMov = rsProd("dUltCierre")
'        nTasa = rsProd("nTasaInteres")
'        nIntAcum = rsProd("nIntAcum")
'        nIntPagado = rsProd("nIntPag")
'        nPlazo = rsProd("nPlazo")
'        dVencimiento = DateAdd("d", nPlazo, dApertura)
'        nPlazoCumplido = DateDiff("d", dApertura, dFecSis)
'        Set clsDef = New nCapDefinicion
'        nTasaDias = clsDef.GetCapTasaInteres(gCapPlazoFijo, CLng(Mid(sCuenta, 9, 1)), gCapTasaNormal, nPlazoCumplido, nSaldoCnt, sCodAge)
'
'        Set clsDef = Nothing
'        nInt01 = 0: nInt02 = 0: nInt03 = 0
'        If dVencimiento = dFecSis Then
'            nInt01 = nIntAcum
'            nDias = DateDiff("d", dUltMov, dVencimiento) - 1
'            nInt02 = GetInteres(nSaldoDisp + nIntAcum, nTasa, nDias, TpoCalcIntCompuesto)
'            nTasaCancelacion = nTasa
'        ElseIf dVencimiento > dFecSis Then
'            Dim sValor As String
'            Set clsCap = New DCapMovimientos
'            sValor = clsCap.GetFechaValorChqApePF(sCuenta)
'            Set clsCap = Nothing
'            If sValor <> "" Then
'                dApertura = CDate(sValor)
'            End If
'            nPlazoCumplido = DateDiff("d", dApertura, dFecSis)
'            nInt01 = GetInteres(nSaldoDisp, nTasaDias, nPlazoCumplido, TpoCalcIntCompuesto) - nIntPagado
'            nTasaCancelacion = nTasaDias
'        Else
'            nInt01 = nIntAcum
'            If dUltMov < dVencimiento Then
'                nInt02 = GetInteres(nSaldoDisp + nInt01, nTasa, DateDiff("d", dUltMov, dVencimiento) - 1, TpoCalcIntCompuesto)
'            Else
'                nInt02 = 0
'            End If
'            nTasaCancelacion = nTasa
'            nInt03 = GetInteres(nSaldoDisp + nInt01 + nInt02, nTasa, DateDiff("d", dVencimiento, dFecSis) - 1, TpoCalcIntCompuesto)
'        End If
'        nIntGanado = nInt01 + nInt02 + nInt03
'        GetSaldoCancelacion = nSaldoDisp + nIntGanado
'        rsProd.Close
'        Set rsProd = Nothing
'End Select
'End Function

'Public Function TieneChequesValorizacion(ByVal sCuenta As String) As Boolean
'Dim clsCap As DCapMovimientos
'Set clsCap = New DCapMovimientos
'TieneChequesValorizacion = clsCap.TieneChequesValorizacion(sCuenta)
'Set clsCap = Nothing
'End Function
'
'Public Sub CapCancelaCuentaAho(ByVal sCuenta As String, ByVal sMovNro As String, _
'        ByVal sGlosa As String, Optional sNroDoc As String = "", _
'        Optional nDocumento As TpoDoc = TpoDocNotaCargo, _
'        Optional nEstado As CaptacEstado = gCapEstCancelada, Optional sNomAge As String = "", _
'        Optional sLpt As String = "", Optional sPersLavDinero As String = "")
'
'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim rsCta As Recordset
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
'Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
'Dim nMonto As Double
'Dim dUltMov As Date
'Dim bInactiva As Boolean, bTrans As Boolean, bInmovilizada As Boolean
'Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
'Dim nExtracto As Long, nMovNro As Long
'Dim sMsgOpe As String, sCodUser As String
'Dim nOperacion As CaptacOperacion
'Dim dFecSis As Date
'Dim oCon As NConstSistemas
'Set oCon = New NConstSistemas
'
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'bTrans = False
''Obtiene los datos para el calculo
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
'Set clsMant = Nothing
'nEstado = rsCta("nPrdEstado")
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'dUltMov = rsCta("dUltCierre")
'nTasa = rsCta("nTasaInteres")
'nIntAcum = rsCta("nIntAcum")
'bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
'bInmovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
'
'nExtracto = rsCta("nTransacc") + 1
'rsCta.Close
'Set rsCta = Nothing
''Inicia la transaccion
'Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba
'
'clsCap.dbCmact.BeginTrans
'bTrans = True
'nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'nMonto = nSaldoCnt + nIntAcum + nIntGanado
'dUltMov = DateAdd("d", -1, dFecSis)
'If sNroDoc <> "" Then
'    nOperacion = gAhoCancTransfAct
'    If bInactiva Then nOperacion = gAhoCancTransfInact
'    sMsgOpe = "Cancelación Transferencia"
'    clsCap.ActualizadUltContactoAho sCuenta, dFecSis
'Else
'    nOperacion = gAhoCancAct
'    If bInactiva Then nOperacion = gAhoCancInact
'    sMsgOpe = "Cancelación Efectivo"
'End If
'If sGlosa = "" Then sGlosa = "Cancelación Cuenta = " & sCuenta
'clsCap.ActualizaCargoCaptacion sCuenta, 0, 0, nIntAcum + nIntGanado, dUltMov, sMovNro, True
'clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntAcum + nIntGanado, 0, 0
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntAcum + nIntGanado
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
'clsCap.dbCmact.CommitTrans
'
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "CANCELACION AHORROS"
'Dim clsNMant As NCapMantenimiento
'Set clsNMant = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsNMant.GetNombreTitulares(sCuenta))
'Set clsNMant = Nothing
'Do
'    If sNroDoc <> "" Then
'        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, sNroDoc, nSaldoDisp, Trim(0), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt
'    Else
'        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt
'    End If
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'Exit Sub
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'End Sub

'Public Sub CapCancelaCuentaCTS(ByVal sCuenta As String, ByVal sMovNro As String, _
'            ByVal sGlosa As String, ByVal nOperacion As CaptacOperacion, Optional sNroDoc As String = "", _
'            Optional nDocumento As TpoDoc = TpoDocNotaCargo, Optional sNomAge As String = "", _
'            Optional sLpt As String = "", Optional sPersLavDinero As String = "")
'
'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim rsCta As Recordset
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
'Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
'Dim nMonto As Double
'Dim dUltMov As Date
'Dim bTrans As Boolean
'Dim nEstado As CaptacEstado
'Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
'Dim nExtracto As Long, nMovNro As Long
'Dim sMsgOpe As String, sCodUser As String
'Dim dFecSis As Date
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'bTrans = False
''Obtiene los datos para el calculo
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
'Set clsMant = Nothing
'nEstado = rsCta("nPrdEstado")
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'dUltMov = rsCta("dUltCierre")
'nTasa = rsCta("nTasaInteres")
'nIntAcum = rsCta("nIntAcum")
'nExtracto = rsCta("nTransacc") + 1
'rsCta.Close
'Set rsCta = Nothing
''Inicia la transaccion
'Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba
'clsCap.dbCmact.BeginTrans
'bTrans = True
'nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'nMonto = nSaldoCnt + nIntAcum + nIntGanado
'dUltMov = DateAdd("d", -1, dFecSis)
'If nOperacion = gCTSCancEfec Then
'    sMsgOpe = "Cancelación Efectivo"
'ElseIf nOperacion = gCTSCancTransf Then
'    sMsgOpe = "Cancelación Transferencia"
'End If
'If sGlosa = "" Then
'    sGlosa = "Cancelación Cuenta = " & sCuenta
'Else
'    sGlosa = sGlosa & Chr$(13) & ".Cancelación Cuenta = " & sCuenta
'End If
'clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntAcum + nIntGanado, 0, 0
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntAcum + nIntGanado
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
'clsCap.dbCmact.CommitTrans
'
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'Dim clsImp As NCapMantenimiento
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "CANCELACION CTS"
'Set clsImp = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsImp.GetNombreTitulares(sCuenta))
'Set clsImp = Nothing
'Do
'    If sNroDoc <> "" Then
'        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, sNroDoc, nSaldoDisp, Trim(0), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt
'    Else
'        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt
'    End If
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'Exit Sub
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'End Sub

'Public Function CapTransferenciaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
'            ByVal sMovNro As String, ByVal rsCtaAbo As Recordset, ByVal sGlosa As String, _
'            Optional sNomAge As String = "", Optional sLpt As String = "", _
'            Optional sPersLavDinero As String = "") As Boolean
'
'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim rsCta As Recordset
'Dim nEstado As CaptacEstado
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
'Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double, nMontoAbo As Double
'Dim dUltMov As Date
'Dim nSaldoInac As Double
'Dim bInactiva As Boolean, bTrans As Boolean, bActivaCta As Boolean, bInmovilizada As Boolean
'Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
'Dim nExtracto As Long, nMovNro As Long
'Dim sMsgOpe As String, sCtaAbo As String, sCodUser As String
'Dim nMoneda As Moneda
'Dim nOperacion As CaptacOperacion
'Dim dFecSis As Date
'Dim oCon As NConstSistemas
'Set oCon = New NConstSistemas
'
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'bTrans = False
'bActivaCta = False
'nMoneda = CLng(Mid(sCuenta, 9, 1))
''Obtiene los datos para el calculo
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
'Set clsMant = Nothing
'nEstado = rsCta("nPrdEstado")
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'dUltMov = rsCta("dUltCierre")
'nTasa = rsCta("nTasaInteres")
'nIntAcum = rsCta("nIntAcum")
'bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
'bInmovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
'
'nExtracto = rsCta("nTransacc") + 1
'rsCta.Close
'Set rsCta = Nothing
'
''Inicia la transaccion
'Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba
'clsCap.dbCmact.BeginTrans
'bTrans = True
'If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
'    clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    MsgBox "Cuenta NO Posee SALDO SUFICIENTE", vbInformation, "Aviso"
'    CapTransferenciaAho = False
'    Exit Function
'End If
'sGlosa = sGlosa & "Cargo Transferencia Cuenta = " & sCuenta
'nOperacion = gAhoTransCargo
'If bInactiva Then
'    nSaldoInac = nSaldoCnt
'End If
'nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'dUltMov = DateAdd("d", -1, dFecSis)
'nSaldoCnt = nSaldoCnt - nMonto
'clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
'nSaldoDisp = nSaldoDisp - nMonto
'clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
'sMsgOpe = "Cargo Transferencia"
'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
'clsCap.ActualizadUltContactoAho sCuenta, dFecSis
'
'If bInmovilizada Then
'    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
'        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
'        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
'        clsCap.ActualizaInmovilizadaAct sCuenta
'    End If
'End If
'
'If bInactiva And Not bInmovilizada Then
'    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
'    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
'    clsCap.ActualizaInactivaAct sCuenta
'End If
'
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'Dim sBoletas As String
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "RETIRO AHORROS"
'Dim clsTit As NCapMantenimiento
'Dim oImp As NCapImpBoleta
'Set clsTit = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
'Set clsMant = Nothing
'Set oImp = New NCapImpBoleta
'sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
''Inicio de los abonos de cuenta
'Do While Not rsCtaAbo.EOF
'    sCtaAbo = rsCtaAbo("Cuenta")
'    If CLng(Mid(sCtaAbo, 9, 1)) = gMonedaNacional Then
'        nMontoAbo = rsCtaAbo("Monto S/.")
'        nOperacion = IIf(nMoneda = gMonedaNacional, gAhoTransAbonoMNMN, gAhoTransAbonoMEMN)
'    Else
'        nMontoAbo = rsCtaAbo("Monto $")
'        nOperacion = IIf(nMoneda = gMonedaNacional, gAhoTransAbonoMNME, gAhoTransAbonoMEME)
'    End If
'    Set clsMant = New DCapMantenimiento
'    Set rsCta = clsMant.GetDatosCuentaAho(sCtaAbo)
'    Set clsMant = Nothing
'    nEstado = rsCta("nPrdEstado")
'    nSaldoDisp = rsCta("nSaldoDisp")
'    nSaldoCnt = rsCta("nSaldo")
'    dUltMov = rsCta("dUltCierre")
'    nTasa = rsCta("nTasaInteres")
'    nIntAcum = rsCta("nIntAcum")
'    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
'    bInmovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
'
'    nExtracto = rsCta("nTransacc") + 1
'    rsCta.Close
'    Set rsCta = Nothing
'    If bInactiva Then
'        nSaldoInac = nSaldoCnt
'    End If
'    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'    dUltMov = DateAdd("d", -1, dFecSis)
'    nSaldoCnt = nSaldoCnt + nMontoAbo
'    clsCap.ActualizaSaldoAnteriorAho sCtaAbo, nSaldoDisp
'    nSaldoDisp = nSaldoDisp + nMontoAbo
'    clsCap.ActualizaAbonoCaptacion sCtaAbo, nMontoAbo, nMontoAbo, nIntGanado, dUltMov, sMovNro, True
'    sMsgOpe = "Abono Transferencia"
'    If sGlosa = "" Then sGlosa = "Abono Transferencia Cuenta = " & sCtaAbo
'    clsCap.AgregaMovCap nMovNro, nOperacion, sCtaAbo, nMontoAbo, nSaldoDisp, nSaldoCnt
'    clsCap.ActualizadUltContactoAho sCtaAbo, dFecSis
'
'    If bInmovilizada Then
'        If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
'            clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCtaAbo, gConcCapital, nSaldoInac
'            clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovNro
'            clsCap.ActualizaInmovilizadaAct sCtaAbo
'        End If
'    End If
'
'    If bInactiva And Not bInmovilizada Then
'        clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCtaAbo, gConcCapital, nSaldoInac
'        clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovNro
'        clsCap.ActualizaInactivaAct sCtaAbo
'    End If
'
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCtaAbo, gConcCapital, nMontoAbo
'    clsCap.UltimaActualizacionCuenta sCtaAbo, sMovNro
'    'Impresión de la boleta
'    sTipDep = IIf(Mid(sCtaAbo, 9, 1) = "1", "SOLES", "DOLARES")
'    sCodOpe = Trim(nOperacion)
'    sModDep = sMsgOpe
'    sTipApe = "TRANSFERENCIA AHORROS"
'    Set clsTit = New NCapMantenimiento
'    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCtaAbo))
'    Set clsTit = Nothing
'    sBoletas = sBoletas & Chr$(12)
'    sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
'    rsCtaAbo.MoveNext
'Loop
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
'clsCap.dbCmact.CommitTrans
'CapTransferenciaAho = True
'Set clsCap = Nothing
'Set oImp = Nothing
'Dim oPrevio As Previo.clsPrevio
'Set oPrevio = New Previo.clsPrevio
'oPrevio.Show sBoletas, "Transferencia de Cuenta"
'Set oPrevio = Nothing
'Exit Function
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsMant = Nothing
'    Set oImp = Nothing
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'    CapTransferenciaAho = False
'End Function

Public Function GetCapInteresAlaFecha(ByVal nProducto As Producto, ByVal nMonto As Double, _
        ByVal nTasa As Double, ByVal dUltMov As Date, ByVal nIntAcum As Double, _
        ByVal dFecSis As Date) As Double
Dim nIntGanado As Double
Dim nDias As Long
nDias = DateDiff("d", dUltMov, dFecSis) - 1
Select Case nProducto
    Case gCapAhorros, gCapCTS
        GetCapInteresAlaFecha = nIntAcum + GetInteres(nMonto, nTasa, nDias, TpoCalcIntSimple)
    Case gCapPlazoFijo
        GetCapInteresAlaFecha = nIntAcum + GetInteres(nMonto, nTasa, nDias, TpoCalcIntCompuesto)
End Select
End Function

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
Set GetMovExtorno = clsCap.GetMovExtorno(sDatoBus, dFecha, nOperacion, nTipoBus)
Set clsCap = Nothing
End Function

Public Function GetMovExtornoCMAC(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
Set GetMovExtornoCMAC = clsCap.GetMovExtornoCMAC(sDatoBus, dFecha, nOperacion, nTipoBus)
Set clsCap = Nothing
End Function

Public Function CapExtornoAbonoAho(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String, Optional sLpt As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
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
bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
    Set clsCap = Nothing
    Err.Raise 1000, "CapExtornoAbonoAho", "Cuenta NO Posee Saldo Suficiente"
    CapExtornoAbonoAho = 0
    Exit Function
End If
sGlosa = sGlosa & "Extorno Abono Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto
clsCap.dbCmact.BeginTrans
bTrans = True
clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocCheque Then
        nSaldoDisp = nSaldoDisp - 0
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
        sMsgOpe = "Ext. Depósito Chq."
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        nSaldoDisp = nSaldoDisp - nMonto
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Ext. Depósito Efectivo"
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoAbonoAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoAbonoAho = 0
End Function

Public Function CapExtornoApertura(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional bDocumento As Boolean = False, Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgProd As String, sCodUser As String
Dim nProd As Producto
Dim nMovNro As Long
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
nProd = CLng(Mid(sCuenta, 6, 3))
If nProd = gCapAhorros Then
    sMsgProd = "EXTORNO AHORROS"
ElseIf nProd = gCapPlazoFijo Then
    sMsgProd = "EXTORNO PLAZO FIJO"
ElseIf nProd = gCapCTS Then
    sMsgProd = "EXTORNO CTS"
End If
Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstAnulada, dFecSis, sMovNro
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, 0, 0
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
If bDocumento Then
    If nTipoDoc = TpoDocCheque Then
        clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
    End If
End If
clsCap.dbCmact.CommitTrans

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = "Extorno Apertura"
sTipApe = sMsgProd
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If bDocumento Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nMonto, 0, "", 1, nMonto, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nMonto, 0, "", 1, nMonto, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
End Function

Public Function CapExtornoCargoAho(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
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
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then sGlosa = "Extorno Cargo Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocOrdenPago Then
        If nOperacion = gAhoExtRetOPCanje Then
            sMsgOpe = "Ext. Retiro OP Canje"
        Else
            sMsgOpe = "Ext. Retiro OP"
        End If
        clsCap.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoCargoAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing

Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCargoAho", Err.Description
    CapExtornoCargoAho = 0
End Function

Public Function CapExtornoCargoAhoOPCertificada(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            ByVal sNroDoc As String, Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset, rsPers As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sNombreIF As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nIntAcum = rsCta("nIntAcum")
nExtracto = rsCta("nTransacc")
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
clsCap.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
If nOperacion = gAhoRetOPCert Then
    sMsgOpe = "Ext.Ret.OP Cert."
ElseIf nOperacion = gAhoRetOPCertCanje Then
    sMsgOpe = "Ext.Ret.OP Cert.Canje"
End If
If sGlosa <> "" Then
    sGlosa = sGlosa & ". Extorno OP Certificada Cuenta = " & sCuenta
Else
    sGlosa = "Extorno OP Certificada Cuenta = " & sCuenta
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoCargoAhoOPCertificada = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing

Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoCargoAhoOPCertificada = 0
End Function

Public Function CapExtornoCargoCTS(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoRetiro As Double, nIntGanadoRet As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long, nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
'APRI20200415 POR COVID-19
Dim nDU01 As Double
Dim oGen As DGeneral
Dim bCeseLaboral As Boolean
Dim nTpoPrograma As Integer
Dim nPorDisp As Double
Dim nSaldRetiro As Double ' JUCS
'END APRI
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nSaldoRetiro = rsCta("nSaldRetiro")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nIntSaldo = rsCta("nIntSaldo")
nExtracto = rsCta("nTransacc") + 1
'APRI20200415 POR COVID-19
nDU01 = rsCta("nDU01")
nPorDisp = rsCta("nPorDisp")
nTpoPrograma = rsCta("nTpoPrograma")
bCeseLaboral = rsCta("bCeseLaboral")
'END APRI
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = "Extorno Cargo Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & ".Extorno Cargo Cuenta = " & sCuenta
End If
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanadoRet = GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
'nSaldoRetiro = nSaldoRetiro + nMonto
'APRI20200415 POR COVID-19
Dim nMontoMN As Double, nTotalSaldos As Double, nTotalSueldos As Double
Dim nExcedente As Double
Dim rsCtaSueldos As ADODB.Recordset
Set rsCtaSueldos = New ADODB.Recordset
Set rsCtaSueldos = clsCap.ObtenerCapSaldosCuentasCTS(sCuenta, oGen.GetTipCambio(dFecSis, TCFijoMes))
nTotalSaldos = rsCtaSueldos!TotalSaldos
nTotalSueldos = rsCtaSueldos!TotalSueldos
Set rsCtaSueldos = Nothing

nMontoMN = nMonto * IIf(Mid(sCuenta, 9, 1) = "1", 1, oGen.GetTipCambio(dFecSis, TCFijoMes))
nExcedente = (nTotalSaldos + IIf(nTpoPrograma = 2 And bCeseLaboral, 0, nMontoMN)) - nTotalSueldos

If nExcedente > 0 Then
   nSaldRetiro = nExcedente * (nPorDisp / 100)
Else
   nSaldRetiro = 0
End If

 nDU01 = nDU01 + IIf((nMonto - nSaldRetiro) < 0, 0, (nMonto - nSaldRetiro))
'END APRI
clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
'clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo + nIntGanadoRet
clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo + nIntGanadoRet, nDU01 'APRI20200415 POR COVID-19
If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoCargoCTS = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO CTS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCargoAho", Err.Description
    CapExtornoCargoCTS = 0
End Function

Public Function CapExtornoCancelacion(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim clsMant As NCapMantenimiento
Dim clsCap As DCapMovimientos
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim dFecSis As Date
Dim bTrans As Boolean
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nExtracto As Long

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nEstado = rsCta("nPrdEstado")
nIntAcum = rsCta("nIntAcum")
nExtracto = rsCta("nTransacc")
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = "Extorno Cancelacion Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & "Extorno Cancelación Cuenta = " & sCuenta
End If
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
If sNroDoc = "" Then
    sMsgOpe = "Ext. Cancelación Efectivo"
Else
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Cancelación NC"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoCancelacion = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCancelacionAho", Err.Description
    CapExtornoCancelacion = 0
End Function

Public Sub CapExtornoTransferenciaAho(ByVal sMovNroBus As String, ByVal sMovNro As String, _
            ByVal sGlosa As String, Optional sNomAge As String = "", Optional sLpt As String = "")
Dim clsCap As DCapMovimientos
Dim rsOpe As Recordset
Dim nOperacion As CaptacOperacion, nOpeExt As CaptacOperacion
Dim nSaldo As Double
Dim sCuenta As String
Dim nMonto As Double
Dim nMovNro As Long
On Error GoTo ErrGraba
Set clsCap = New DCapMovimientos
Set rsOpe = clsCap.GetMovTransferencia(sMovNroBus)
Set clsCap = Nothing
If Not (rsOpe.EOF And rsOpe.BOF) Then
    Do While Not rsOpe.EOF
        nOperacion = CLng(rsOpe("cOpeCod"))
        nMonto = rsOpe("nMonto")
        sCuenta = rsOpe("cCtaCod")
        nMovNro = rsOpe("nMovNro")
        If nOperacion = gAhoTransCargo Then
            nSaldo = CapExtornoCargoAho(nMovNro, nOperacion, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt)
        Else
            nSaldo = CapExtornoAbonoAho(nMovNro, nOperacion, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt)
        End If
        rsOpe.MoveNext
    Loop
End If
ErrGraba:
    Err.Raise Err.Number, "CapExtornoTransferenciaAho", Err.Description
End Sub

Public Function CapRetiroInteresPF(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nMonto As Double, ByVal nDias As Long, ByVal sGlosa As String, _
        ByVal nOperacion As CaptacOperacion, Optional sCuentaAho As String = "", _
        Optional nTipoCambio As Double, Optional sCodCMAC As String = "", _
        Optional sNomCMAC As String, Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date, dAuxiliar As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long, nDiasAux As Long
Dim nExtracto As Long, nPlazo As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
'Obtiene los datos para el calculo del retiro de intereses
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
dAuxiliar = rsCta("dAuxiliar")
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
nDiasAux = 30 * Int(DateDiff("d", dAuxiliar, dFecSis) / 30)
If nOperacion = gPFRetIntAboAho Then
    If sGlosa = "" Then
        sGlosa = "Retiro Intereses PF Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
    Else
        sGlosa = sGlosa & Chr$(13) & "Retiro Intereses PF Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
    End If
ElseIf nOperacion = gPFRetInt Then
    If sGlosa = "" Then
        sGlosa = "Retiro Intereses Efectivo PF Cuenta = " & sCuenta
    Else
        sGlosa = sGlosa & Chr$(13) & "Retiro Intereses Efectivo PF Cuenta = " & sCuenta
    End If
End If
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp + nIntAcum, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
dUltMov = DateAdd("d", -1, dFecSis)
dAuxiliar = DateAdd("d", nDiasAux, dAuxiliar)
clsCap.ActualizaRetiroInteresPF sCuenta, nMonto, nIntGanado, dUltMov, sMovNro, dAuxiliar, True
sMsgOpe = "Retiro Int. PF"
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If nIntGanado > 0 Then
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvision, nIntGanado
End If
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nMonto
If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO PLAZO FIJO"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , Trim(nDias), , , , , , , dFecSis, sNomAge, sCodUser, sLpt
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
bTrans = False

If sPersLavDinero <> "" Then
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
End If

If sCuentaAho <> "" Then
    'Obtiene los datos para el calculo del abono a la cuenta de ahorros
    Dim nEstado As CaptacEstado
    Dim bInactiva As Boolean, bInmovilizada As Boolean
    Dim nSaldoInac As Double
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    Set clsMant = New DCapMantenimiento
    Set rsCta = clsMant.GetDatosCuentaAho(sCuentaAho)
    Set clsMant = Nothing
    nEstado = rsCta("nPrdEstado")
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    dUltMov = rsCta("dUltCierre")
    nTasa = rsCta("nTasaInteres")
    nIntAcum = rsCta("nIntAcum")
    bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
    bInmovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
    
    nExtracto = rsCta("nTransacc") + 1
    rsCta.Close
    Set rsCta = Nothing
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
    If nTipoCambio <> 0 Then
        nMonto = Round(nMonto / nTipoCambio, 2)
    End If
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMonto
    clsCap.ActualizaSaldoAnteriorAho sCuentaAho, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto
    clsCap.ActualizaAbonoCaptacion sCuentaAho, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Depósito Int. PF"
    
    clsCap.AgregaMovCap nMovNro, nOperacion, sCuentaAho, nMonto, nSaldoDisp, nSaldoCnt
    clsCap.ActualizadUltContactoAho sCuentaAho, dFecSis
            
    If bInmovilizada Then
        If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
            clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuentaAho, gConcCapital, nSaldoInac
            clsCap.ActualizaEstadoCuenta sCuentaAho, gCapEstActiva, dFecSis, sMovNro
            clsCap.ActualizaInmovilizadaAct sCuentaAho
        End If
    End If
    
    If bInactiva And Not bInmovilizada Then
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuentaAho, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCuentaAho, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInactivaAct sCuentaAho
    End If
    
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuentaAho, gConcCapital, nMonto
    clsCap.UltimaActualizacionCuenta sCuentaAho, sMovNro
    clsCap.dbCmact.CommitTrans
    
    sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "DEPOSITO AHORROS"
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
    Set clsTit = Nothing
    Do
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sLpt
    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Else
    clsCap.dbCmact.CommitTrans
End If
CapRetiroInteresPF = nSaldoCnt
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapRetiroInteresPF = 0
End Function

Public Function GetMontoRetiroIntPF(ByVal nCapital As Double, ByVal dUltMov As Date, _
        ByVal nTasa As Double, ByVal dFecha As Date, ByVal nIntAcum As Double) As Double
Dim nDiasTranscurridos As Long
Dim nIntGanado  As Double
nDiasTranscurridos = DateDiff("d", dUltMov, dFecha) - 1
nIntGanado = GetInteres(nCapital + nIntAcum, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
GetMontoRetiroIntPF = nIntAcum + nIntGanado
End Function

Public Function GetFechaUltimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
GetFechaUltimoRetiroIntPF = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
Set clsCap = Nothing
End Function

'Public Function CapCancelacionPF(ByVal sCuenta As String, ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
'            ByVal sGlosa As String, ByVal nCapital As Double, ByVal nIntGanado As Double, _
'            Optional sNroDoc As String = "", Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, _
'            Optional sNomAge As String = "", Optional sLpt As String = "", _
'            Optional sPersLavDinero As String = "") As Double
'
'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim rsCta As Recordset
'Dim nEstado As CaptacEstado
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
'Dim nTasa As Double, nSaldoCanc As Double
'Dim dUltMov As Date, dAuxiliar As Date
'Dim bTrans As Boolean
'Dim nDiasTranscurridos As Long, i As Long, nDiasAux As Long
'Dim nExtracto As Long, nPlazo As Long, nMovNro As Long
'Dim sMsgOpe As String, sCodUser As String, sCodAge As String
'Dim dFecSis As Date
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'sCodAge = Mid(sCuenta, 4, 2)
'bTrans = False
''Obtiene los datos para el calculo
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
'Set clsMant = Nothing
'nEstado = rsCta("nPrdEstado")
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'dUltMov = rsCta("dUltCierre")
'nTasa = rsCta("nTasaInteres")
'nIntAcum = rsCta("nIntAcum")
'dAuxiliar = rsCta("dAuxiliar")
'nExtracto = rsCta("nTransacc") + 1
'rsCta.Close
'Set rsCta = Nothing
''Verifica el estado de la cuenta por si acaso, no vaya a haber algun chistoso que haya
''hecho algo con la cuenta
'If nEstado <> gCapEstActiva Then
'    Select Case nEstado = gCapEstCancelada
'        Case gCapEstCancelada, gCapEstAnulada
'            Err.Raise 1000, "CapCancelacionPF", "Cuenta Cancelada o Anulada"
'        Case gCapEstBloqTotal, gCapEstBloqRetiro
'            Err.Raise 1000, "CapCancelacionPF", "Cuenta Bloqueada"
'    End Select
'    CapCancelacionPF = 0
'    Exit Function
'End If
''Verifica si en alguna parte de la Caja se retiraron intereses antes de grabar
'nSaldoCanc = GetSaldoCancelacion(sCuenta, dFecSis, sCodAge)
'If Round(nSaldoCanc, 2) <> Round(nSaldoDisp + nIntGanado, 2) Then
'    Err.Raise 1000, "CapCancelacionPF", "Se retiraron intereses antes de efectuar la cancelación. " & Chr$(13) _
'        & "Por favor consulte nuevamente el saldo de la cuenta"
'    CapCancelacionPF = 0
'    Exit Function
'End If
''Inicia la transaccion
'Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba
'clsCap.dbCmact.BeginTrans
'If sGlosa = "" Then
'    sGlosa = "Cancelación PF Cuenta = " & sCuenta
'Else
'    sGlosa = sGlosa & Chr$(13) & "Cancelación PF Cuenta = " & sCuenta
'End If
'clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
'nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntGanado, 0, 0
'If nIntGanado - nIntAcum > 0 Then
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvision, nIntGanado - nIntAcum
'Else
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcRevProvision, Abs(nIntGanado - nIntAcum)
'End If
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
'If nIntGanado < 0 Then
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvisionNeg, Abs(nIntGanado)
'ElseIf nIntGanado > 0 Then
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntGanado
'End If
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
'clsCap.dbCmact.CommitTrans
'CapCancelacionPF = nSaldoCnt
'
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "CANCELACION PLAZO FIJO"
'Dim clsTit As NCapMantenimiento
'Set clsTit = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
'Set clsTit = Nothing
'Do
'    If sNroDoc <> "" Then
'        ImprimeBoletaInteres sTipApe, "Capital", sCodOpe, Trim(nSaldoDisp), sNomTit, sCuenta, sNroDoc, nSaldoDisp, Trim(0), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntGanado), , , , sNomAge, dFecSis, sCodUser
'    Else
'        ImprimeBoletaInteres sTipApe, "Capital", sCodOpe, Trim(nSaldoDisp), sNomTit, sCuenta, "", nSaldoDisp, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntGanado), , , , sNomAge, dFecSis, sCodUser
'    End If
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'Exit Function
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'    CapCancelacionPF = 0
'End Function
'
'Public Sub EmiteDuplicadoCertPF(ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String)
'Dim clsCap As DCapMovimientos
'Dim clsMant As DCapMantenimiento
'Dim nMovNro As Long
'Dim bTrans As Boolean
'Dim nSaldoDisp As Double, nSaldoCnt As Double
'Dim rsCta As Recordset
'
'bTrans = False
''Obtiene los datos de la Cuenta
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
'Set clsMant = Nothing
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'On Error GoTo ErrGraba
'Set clsCap = New DCapMovimientos
'clsCap.dbCmact.BeginTrans
'bTrans = True
'If sGlosa = "" Then
'    sGlosa = "Emisión Duplicado de Plazo Fijo Cuenta = " & sCuenta
'Else
'    sGlosa = sGlosa & Chr$(13) & "Emisión Duplicado de Plazo Fijo Cuenta = " & sCuenta
'End If
'clsCap.AgregaMov sMovNro, gPFDupCert, sGlosa, gMovEstContabNoContable
'nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, gPFDupCert, sCuenta, 0, nSaldoDisp, nSaldoCnt
'clsCap.ActualizaDuplicadoCertPF sCuenta
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'clsCap.dbCmact.CommitTrans
'bTrans = False
'Exit Sub
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'End Sub
'
'Public Function CapExtornoRetiroIntPF(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
'            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
'            Optional sNomAge As String = "", Optional sLpt As String = "") As Double
'
'Dim nMovNro As Long, nExtracto As Long, i As Long
'Dim sMsgOpe As String
'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntGanado As Double
'Dim nTasa As Double, nIntAcum As Double
'Dim dAuxiliarAnt As Date, dUltMovAnt As Date
'Dim dUltMov As Date
'Dim nEstado As CaptacEstado
'Dim bTrans As Boolean
'Dim rsCta As Recordset
'Dim nSaldoInac As Double, nMontoAho As Double
'Dim bInactiva As Boolean
'Dim sCuentaAho As String, sCodUser As String
'Dim nDiasTranscurridos As Long
'Dim dFecSis As Date
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'Dim clsTit As NCapMantenimiento
'
'dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'sCodUser = Right(sMovNro, 4)
'bTrans = False
'
''Inicia la transaccion
'Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba
'clsCap.dbCmact.BeginTrans
'bTrans = True
''Actualiza y agrega operacion de extorno
'If sGlosa = "" Then
'    If nOperacion = gPFRetIntAboAho Then
'        sGlosa = "Extorno Retiro Interes PF. Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
'    Else
'        sGlosa = "Extorno Retiro Interes PF. Cuenta = " & sCuenta
'    End If
'Else
'    If nOperacion = gPFRetIntAboAho Then
'        sGlosa = sGlosa & Chr$(13) & "Extorno Retiro Interes PF. Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
'    Else
'        sGlosa = sGlosa & Chr$(13) & "Extorno Retiro Interes PF. Cuenta = " & sCuenta
'    End If
'End If
'clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
'clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
'nMovNro = clsCap.GetnMovNro(sMovNro)
'
'If nOperacion = gPFExtRetIntAboAho Then 'Si el extorno es con abono a cuenta de ahorros
'    Set rsCta = clsCap.GetCapMov(nMovNroBus, gConcCapital, gCapAhorros)
'    sCuentaAho = rsCta("cCtaCod")
'    nMontoAho = rsCta("nMonto")
'    'Obtiene los datos para el calculo
'    Set clsMant = New DCapMantenimiento
'    Set rsCta = clsMant.GetDatosCuentaAho(sCuentaAho)
'    Set clsMant = Nothing
'    nEstado = rsCta("nPrdEstado")
'    nSaldoDisp = rsCta("nSaldoDisp")
'    nSaldoCnt = rsCta("nSaldo")
'    dUltMov = rsCta("dUltCierre")
'    nTasa = rsCta("nTasaInteres")
'    nIntAcum = rsCta("nIntAcum")
'    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
'    nExtracto = rsCta("nTransacc") + 1
'    rsCta.Close
'    Set rsCta = Nothing
'    Randomize
'    For i = 0 To Rnd(2000) * 1000
'    Next i
'    If Not ValidaSaldoCuenta(sCuentaAho, nMontoAho) Then
'        clsCap.dbCmact.RollbackTrans
'        Set clsCap = Nothing
'        Err.Raise 1000, "CapExtornoRetiroIntPF", "Cuenta NO Posee Saldo Suficiente"
'        CapExtornoRetiroIntPF = 0
'        Exit Function
'    End If
'    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'    dUltMov = DateAdd("d", -1, dFecSis)
'    nSaldoCnt = nSaldoCnt - nMonto
'    clsCap.ActualizaSaldoAnteriorAho sCuentaAho, nSaldoDisp
'    nSaldoDisp = nSaldoDisp - nMonto
'    clsCap.ActualizaCargoCaptacion sCuentaAho, nMontoAho, nMontoAho, nIntGanado, dUltMov, sMovNro, True
'    clsCap.AgregaMovCap nMovNro, nOperacion, sCuentaAho, nMontoAho, nSaldoDisp, nSaldoCnt
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuentaAho, gConcCapital, nMontoAho
'    clsCap.UltimaActualizacionCuenta sCuentaAho, sMovNro
'    sMsgOpe = "Ext.Ret.Int.PF Abo.Cta.Aho"
'
'    sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'    sCodOpe = Trim(nOperacion)
'    sModDep = sMsgOpe
'    sTipApe = "EXTORNO AHORROS"
'
'    Set clsTit = New NCapMantenimiento
'    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuentaAho))
'    Set clsTit = Nothing
'    Do
'        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAho), sNomTit, sCuentaAho, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt
'    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'
'Else
'    sMsgOpe = "Ext. Retiro Interes PF"
'End If
'
''Obtiene los datos para el cálculo
'Set clsMant = New DCapMantenimiento
'Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
'Set clsMant = Nothing
'nSaldoDisp = rsCta("nSaldoDisp")
'nSaldoCnt = rsCta("nSaldo")
'nEstado = rsCta("nPrdEstado")
'dAuxiliarAnt = rsCta("dAuxiliarAnt")
'dUltMovAnt = rsCta("dUltCierreAnt")
'nExtracto = rsCta("nTransacc")
'rsCta.Close
'Set rsCta = Nothing
'Set rsCta = clsCap.GetCapMov(nMovNroBus, gConcProvision, gCapPlazoFijo)
'nIntGanado = rsCta("nMontoConcepto")
'
'clsCap.ActualizaExtornoRetiroInteresPF sCuenta, nMonto, nIntGanado, dUltMovAnt, sMovNro, dAuxiliarAnt
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nMonto
'clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'clsCap.dbCmact.CommitTrans
'CapExtornoRetiroIntPF = nSaldoCnt
'
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "EXTORNO AHORROS"
'Set clsTit = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
'Set clsTit = Nothing
'Do
'    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'
'Exit Function
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "CapExtornoCancelacionAho", Err.Description
'    CapExtornoRetiroIntPF = 0
'End Function
'

Public Sub CapConsultaSaldosMovimiento(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nOperacion As CaptacOperacion)
Dim clsCap As DCapMovimientos
Dim clsMant As NCapMantenimiento
Dim nMovNro As Long
Dim bTrans As Boolean
Dim nSaldoDisp As Double, nSaldoCnt As Double
Dim rsCta As Recordset
Dim sGlosa As String
bTrans = False
'Obtiene los datos de la Cuenta
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
On Error GoTo ErrGraba
Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
bTrans = True
Select Case nOperacion
    Case gAhoConsSaldo, gPFConsSaldo, gCTSConsSaldo
        sGlosa = "Consulta de Saldo Cuenta = " & sCuenta
    Case gAhoConsMovimiento, gPFConsMovimiento, gCTSConsMovimiento
        sGlosa = "Consulta Movimientos Cuenta = " & sCuenta
End Select
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, 0, nSaldoDisp, nSaldoCnt
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Function GetCuentaAbonoIntPF(ByVal sCuenta As String) As String
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
GetCuentaAbonoIntPF = clsCap.GetCuentaAbonoIntPF(sCuenta)
Set clsCap = Nothing
End Function

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
EsOrdenPagoEmitida = clsCap.EsOrdenPagoEmitida(sCuenta, nNumOP)
Set clsCap = Nothing
End Function

Public Function CapExtornoAbonoCTS(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", _
            Optional sLpt As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nPorcDisp As Double, nSaldRetiro As Double, nIntGanadoRet As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim nDU01 As Double 'APRI20200415 POR COVID-19
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nSaldRetiro = rsCta("nSaldRetiro")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nIntSaldo = rsCta("nIntSaldo")
nExtracto = rsCta("nTransacc") + 1
nDU01 = rsCta("nDU01") 'APRI20200415 POR COVID-19
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion
Set clsCap = New DCapMovimientos
nPorcDisp = clsCap.GetCapMovParamCTS(nMovNroBus)
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
    clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise 1000, "CapExtornoAbonoCTS", "Cuenta NO Posee Saldo Suficiente"
    CapExtornoAbonoCTS = 0
    Exit Function
End If
If sGlosa = "" Then
    sGlosa = "Extorno Abono CTS Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & "Extorno Abono CTS Cuenta = " & sCuenta
End If
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanadoRet = GetInteres(nSaldRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocCheque Then
        nSaldoDisp = nSaldoDisp - 0
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
        sMsgOpe = "Ext. Depósito Chq."
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        nSaldRetiro = nSaldRetiro - (nMonto * nPorcDisp / 100)
        nSaldoDisp = nSaldoDisp - nMonto
        'clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo
        clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldRetiro = nSaldRetiro - (nMonto * nPorcDisp / 100)
    nSaldoDisp = nSaldoDisp - nMonto
    'clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo
    clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Ext. Depósito Efectivo"
End If

clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoAbonoCTS = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO CTS"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoAbonoCTS = 0
End Function

Public Sub CapOpeAhoCMACLlamada(ByVal sMovNro As String, ByVal nMoneda As Moneda, _
            ByVal sGlosa As String, ByVal nOperacion As CaptacOperacion, ByVal nExtracto As Integer, _
            ByVal sDescOperacion As String, ByVal nMonto As Double, ByVal sCuenta As String, _
            ByVal nSaldo As Double, ByVal sCodCMAC As String, ByVal sNomCMAC As String, _
            ByVal sCliente As String, Optional nTpoDoc As TpoDoc = TpoDocOrdenPago, _
            Optional sDocumento As String = "NULL", Optional sCuentaAho As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = sDescOperacion & ", " & sNomCMAC & ", Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & sDescOperacion & ", " & sNomCMAC & ", Cuenta = " & sCuenta
End If
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCMAC nMovNro, sCodCMAC, nMoneda, nOperacion, nMonto, sDocumento, sCuenta
nMovNro = clsCap.GetnMovNro(sMovNro)

If sCuentaAho <> "" Then
    Select Case nOperacion
        Case gCMACOTAhoDepEfec, gCMACOTAhoDepChq
            CapAbonoCuentaAho sCuentaAho, nMonto, nOperacion, sMovNro, sGlosa, , , , , , , , , sNomAge, sLpt, , True
        Case gCMACOTAhoRetEfec, gCMACOTAhoRetOP
            CapCargoCuentaAho sCuentaAho, nMonto, nOperacion, sMovNro, sGlosa, , , , , , , , , sNomAge, sLpt, , True
    End Select
Else
    
End If
clsCap.dbCmact.CommitTrans
If sDocumento = "NULL" Then
    ImprimeBoleta "OPERACION CMAC LLAMADA", sDescOperacion, Trim(nOperacion), Trim(nMonto), sCliente, sCuenta, "", nSaldo, "", "", nExtracto, nSaldo, True, True, , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
Else
    ImprimeBoleta "OPERACION CMAC LLAMADA", sDescOperacion & " N° " & sDocumento, Trim(nOperacion), Trim(nMonto), sCliente, sCuenta, "", nSaldo, "", "", nExtracto, nSaldo, True, True, , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
End If
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub CapExtornoOpeAhoCMACLlamada(ByVal nMovNroBus As Long, ByVal sMovNro As String, _
            ByVal nOperacion As CaptacOperacion, ByVal sGlosa As String, ByVal sCuenta As String, _
            ByVal sDescOperacion As String, ByVal nMonto As Double, _
            Optional sDocumento As String = "NULL", Optional sNomAge As String, _
            Optional sLpt As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
clsCap.dbCmact.CommitTrans
If sDocumento = "" Then
    ImprimeBoleta "EXT.OPE CMAC LLAMADA", sDescOperacion, Trim(nOperacion), Trim(nMonto), "", sCuenta, "", 0, "", "", 0, 0, True, False, , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
Else
    ImprimeBoleta "EXT.OPE CMAC LLAMADA", sDescOperacion & " N° " & sDocumento, Trim(nOperacion), Trim(nMonto), "", sCuenta, "", 0, "", "", 0, 0, True, False, , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
End If
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Sub CapCargoFondoFijo(ByVal sNroDoc As String, ByVal nMonto As Double, ByVal nMovNroRef As Long, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal dFecDoc As Date, _
            ByVal sGlosa As String, Optional sNomAge As String = "", Optional sLpt As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovDoc nMovNro, TpoDocOrdenPago, sNroDoc, dFecDoc
clsCap.AgregaMovRef nMovNro, nMovNroRef
clsCap.dbCmact.CommitTrans
Dim sMensaje As String
Select Case nOperacion
    Case gAhoRetFondoFijo
        sMensaje = "Ord. Pag. N° " & sNroDoc
    Case gAhoRetFondoFijoCanje
        sMensaje = "OP Cert. N° " & sNroDoc
End Select
ImprimeBoleta "RETIRO FONDO FIJO", sMensaje, Trim(nOperacion), Trim(nMonto), "", "", sNroDoc, 0, "", "", 0, 0, False, False, , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollBackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function CtaConFirmas(psCtaCod As String) As Boolean
    Dim oCapMov As DCapMovimientos
    Set oCapMov = New DCapMovimientos
    
    CtaConFirmas = oCapMov.CtaConFirmas(psCtaCod)
    
    Set oCapMov = Nothing
End Function

Public Function OtrasOperaciones(psMovNro As String, pnOperacion As CaptacOperacion, pnMovImporte As Currency, psNroDoc As String, psReferencia As String, pnMoneda As Moneda, psPersCod As String) As Long
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim lnMovNro As Long
    
    clsCapMov.AgregaMov psMovNro, pnOperacion, psReferencia, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = clsCapMov.GetnMovNro(psMovNro)
    clsCapMov.AgregaMovOpeVarias lnMovNro, pnMovImporte, psNroDoc, psReferencia, pnMoneda
    clsCapMov.AgregaMovGasto lnMovNro, psPersCod, ""
    
    Set clsCapMov = Nothing
    OtrasOperaciones = lnMovNro
End Function

Public Function GetTarifa(psOpeCod As CaptacOperacion, pnMoneda As Moneda) As Currency
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    
    GetTarifa = clsCapMov.GetTarifa(psOpeCod, pnMoneda)
    
    Set clsCapMov = Nothing
End Function

Public Function GetOtrasOperaciones(pdFecha As Date) As ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = clsCapMov.GetOtrasOperaciones(pdFecha)
    
    Set GetOtrasOperaciones = rs
    
    Set clsCapMov = Nothing
End Function

Public Function GetOtrasOperacionesDet(psMovNro As Long) As ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    
    Set GetOtrasOperacionesDet = clsCapMov.GetOtrasOperacionesDet(psMovNro)
    
    Set clsCapMov = Nothing
End Function

Public Function OtrasOperacionesExtorno(psMovNro As String, pnMovNro As Long, pnOperacion As CaptacOperacion, psReferencia As String) As Long
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim lnMovNro As Long
    
    clsCapMov.AgregaMov psMovNro, pnOperacion, psReferencia, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = clsCapMov.GetnMovNro(psMovNro)
    clsCapMov.AgregaMovRef lnMovNro, pnMovNro
    clsCapMov.ActualizaEstadoMov pnMovNro, gMovFlagEliminado
    
    Set clsCapMov = Nothing
    OtrasOperacionesExtorno = lnMovNro
End Function

'Public Function GetCreditosPendientes(ByVal sPersona As String, pgdFecSis As Date, Optional ByVal pbBusqCred As Boolean = True, _
'        Optional pbBusqPrend As Boolean = True, Optional pbBusqJud As Boolean = True) As ADODB.Recordset
'    Dim rsCred As ADODB.Recordset
'    Set rsCred = New ADODB.Recordset
'    Dim rsCta As ADODB.Recordset
'    Set rsCta = New ADODB.Recordset
'    Dim clsCapMov As DCapMovimientos
'    Set clsCapMov = New DCapMovimientos
'    Dim nCuentasCre As Integer
'    Dim nCuentasPrd As Integer
'    Dim nCuentasJud As Integer
'
'    nCuentasCre = 0
'    nCuentasPrd = 0
'    nCuentasJud = 0
'
'    GeneraRecordsetCred rsCred
'    rsCred.Open
'
'    ReDim Preserve CuentasPersonaCre(0)
'    If pbBusqCred = True Then
'        Set rsCta = clsCapMov.CargaCuentas(sPersona, "CRE")
'
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasCre = nCuentasCre + 1
'                ReDim Preserve CuentasPersonaCre(nCuentasCre)
'                CuentasPersonaCre(nCuentasCre) = rsCta("cCtaCod")
'                rsCta.MoveNext
'            Loop
'        End If
'    End If
'
'    ReDim Preserve CuentasPersonaPrd(0)
'    If pbBusqPrend = True Then
'        Set rsCta = clsCapMov.CargaCuentas(sPersona, "PRD")
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasPrd = nCuentasPrd + 1
'                ReDim Preserve CuentasPersonaPrd(nCuentasPrd)
'                CuentasPersonaPrd(nCuentasPrd) = rsCta("cCtaCod")
'                rsCta.MoveNext
'            Loop
'        End If
'    End If
'
'    ReDim Preserve CuentasPersonaJud(0)
'    If pbBusqJud = True Then
'        Set rsCta = clsCapMov.CargaCuentas(sPersona, "JUD")
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasJud = nCuentasJud + 1
'                ReDim Preserve CuentasPersonaJud(nCuentasJud)
'                CuentasPersonaJud(nCuentasJud) = rsCta("cCtaCod")
'                rsCta.MoveNext
'            Loop
'        End If
'    End If
'    CreditosPendientes sPersona, rsCred, pgdFecSis
'
'    Set GetCreditosPendientes = rsCred
'End Function

Public Sub GeneraRecordsetCred(ByRef rsCred As ADODB.Recordset)
    rsCred.Fields.Append "Cuenta", adVarChar, 18, adFldMayBeNull
    rsCred.Fields.Append "Estado", adVarChar, 2, adFldMayBeNull
    rsCred.Fields.Append "Vencimiento", adDate, , adFldMayBeNull
    rsCred.Fields.Append "CuoVenc", adInteger, , adFldMayBeNull
    rsCred.Fields.Append "Cuota", adCurrency, , adFldMayBeNull
    rsCred.Fields.Append "Mora", adCurrency, , adFldMayBeNull
    rsCred.Fields.Append "Gastos", adCurrency, , adFldMayBeNull
    rsCred.Fields.Append "MinPagar", adCurrency, , adFldMayBeNull
End Sub

'Public Function CreditosPendientes(ByVal sPersona As String, ByRef rsCred As ADODB.Recordset, pgdFecSis As Date) As Boolean
'    Dim rsCta As ADODB.Recordset
'    Dim nAgencias As Long, nLen As Long, nCuentasCre As Long, nCuentasPrd As Long, nCuentasJud As Long
'    Dim sBuffer As String, sAgencia As String, sTipo As String
'    Dim sSimbMoneda As String, sCuentas As String, sVencidos As String
'    Dim i As Integer, J As Integer, nDiasAtraso As Integer
'    Dim nDeudaPrend As Double, nMinimoPagar As Double
'    Dim bVencidos As Boolean
'    Dim nValor1 As Double, nValor2 As Double
'    Dim clsCapDef As nCapDefinicion
'    Set clsCapDef = New nCapDefinicion
'
'    Dim clsCapMov As DCapMovimientos
'    Set clsCapMov = New DCapMovimientos
'
'    sAgencia = ""
'    sBuffer = ""
'    nCuentasCre = 0
'
'    nValor1 = clsCapDef.GetCapParametro(2027)
'    nValor2 = clsCapDef.GetCapParametro(2028)
'
'    If nValor1 = 1 Then
'        bVencidos = True
'        nDiasAtraso = Int(nValor2)
'    Else
'        bVencidos = False
'    End If
'    Set rsCta = New ADODB.Recordset
'    rsCta.CursorLocation = adUseClient
'
'    For i = 1 To UBound(CuentasPersonaCre, 1)
'        sCuentas = "("
'        sAgencia = Left(CuentasPersonaCre(i), 2)
'        For J = i To UBound(CuentasPersonaCre, 1)
'            sBuffer = Left(CuentasPersonaCre(J), 2)
'            If sBuffer = sAgencia Then
'                sCuentas = sCuentas & "'" & CuentasPersonaCre(J) & "',"
'            Else
'                Exit For
'            End If
'        Next J
'        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'        sCuentas = sCuentas & ")"
'        i = J - 1
'
'        Set rsCta = clsCapMov.GetDatosCrePendiente("CRE", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasCre = nCuentasCre + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
'                rsCred("Estado") = IIf(rsCta("cRefinan") = 1, "RF", "")
'                rsCred("Vencimiento") = CDate(rsCta("dFecVenc"))
'                rsCred("CuoVenc") = Trim(rsCta("nNroCuo"))
'                rsCred("Cuota") = rsCta("nCuota")
'                rsCred("Mora") = rsCta("nMora")
'                rsCred("Gastos") = rsCta("nGastos")
'                rsCred("MinPagar") = rsCta("nCuota") + rsCta("nGastos") + rsCta("nMora")
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    Next i
'
'    'Verifica los Créditos Prendarios
'    sAgencia = ""
'    sBuffer = ""
'    nCuentasPrd = 0
'    For i = 1 To UBound(CuentasPersonaPrd, 1)
'        sCuentas = "("
'        sAgencia = Left(CuentasPersonaPrd(i), 2)
'        For J = i To UBound(CuentasPersonaPrd, 1)
'            sBuffer = Left(CuentasPersonaPrd(J), 2)
'            If sBuffer = sAgencia Then
'                sCuentas = sCuentas & "'" & CuentasPersonaPrd(J) & "',"
'            Else
'                Exit For
'            End If
'        Next J
'        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'        sCuentas = sCuentas & ")"
'        i = J - 1
'
'    Set rsCta = clsCapMov.GetDatosCrePendiente("CRE", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
'    If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasPrd = nCuentasPrd + 1
'                nDeudaPrend = CalculaDeudaPrendario(rsCta("nSaldoCap"), rsCta("dFecVenc"), rsCta("nValTasac"), rsCta("nTasaIntVenc"), _
'                                    rsCta("nCostCusto"), rsCta("nTasaImpuesto"), rsCta("nPRDEstado"), rsCta("nCostoPrepRemate"), pgdFecSis)
'                nMinimoPagar = CalculaMinimoPagar(rsCta("nTasaInt"), rsCta("nPlazo"), rsCta("nValTasac"), rsCta("nTasaCustodia"), _
'                                    rsCta("nSaldoCap"), rsCta("nPRDEstado"), rsCta("dFecVenc"), rsCta("nTasaIntVenc"), rsCta("nTasaCustodia"), _
'                                    rsCta("nTasaImpuesto"), rsCta("nCostoPrepRemate"), pgdFecSis)
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
'                rsCred("Estado") = IIf(rsCta("nPrdEstado") = ColocEstado.gColPEstPRema, "RE", "")
'                rsCred("Vencimiento") = rsCta("dFecVenc")
'                rsCred("CuoVenc") = 1
'                rsCred("Cuota") = rsCta("nSaldoCap")
'                rsCred("Mora") = nDeudaPrend - rsCta("nSaldoCap")
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = nMinimoPagar
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'
'    Next i
'
'    nCuentasJud = 0
'
'    If UBound(CuentasPersonaJud, 1) > 0 Then
'        sCuentas = "("
'        For i = 1 To UBound(CuentasPersonaJud, 1)
'            sCuentas = sCuentas & "'" & CuentasPersonaJud(i) & "',"
'        Next i
'        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'        sCuentas = sCuentas & ")"
'        'Verificamos los Creditos Judiciales concentrados en la agencia principal
'
'        Set rsCta = clsCapMov.GetDatosCrePendiente("PRD", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasJud = nCuentasJud + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
'                rsCred("Estado") = "JU"
'                rsCred("Vencimiento") = Null
'                rsCred("CuoVenc") = 0
'                rsCred("Cuota") = 0
'                rsCred("Mora") = 0
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = 0
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'
'    If (nCuentasCre + nCuentasPrd + nCuentasJud) = 0 Then
'        CreditosPendientes = False
'    Else
'        CreditosPendientes = True
'    End If
'End Function
'
'Public Function BuscaCreditosPendientesPago(ByVal psCuenta As String) As Boolean
'    Dim clsCapMov As DCapMovimientos
'    Set clsCapMov = New DCapMovimientos
'
'    BuscaCreditosPendientesPago = clsCapMov.BuscaCreditosPendientesPago(psCuenta)
'
'    Set clsCapMov = Nothing
'End Function

Public Function GetFirma(psPersCod As String) As ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
        
    Set GetFirma = clsCapMov.GetFirma(psPersCod)

    Set clsCapMov = Nothing
End Function

'Public Function CalculaDeudaPrendario(pnSaldoCapital As Double, _
'    pdFecVencimiento As Date, pnValorTasacion As Double, _
'    pnTasaInteresVencido As Double, pnCostoCustodiaMoratoria As Double, _
'    pnTasaImpuesto As Double, pnEstado As ColocEstado, _
'    pnCostoPreparacionRemate As Double, pgdFecSis As Date, Optional pdFecParaDeuda As String) As Double
'    Dim vDiasAtra As Double
'    Dim vDeud As Double
'    Dim vInteMora As Double
'    Dim vImpu As Double
'    Dim vCostCustMora As Double
'    Dim clsCalculos As NColPCalculos
'    Set clsCalculos = New NColPCalculos
'
'    If Len(pdFecParaDeuda) <= 0 Then pdFecParaDeuda = pgdFecSis
'    vDiasAtra = IIf(DateDiff("d", pdFecVencimiento, pdFecParaDeuda) <= 0, 0, DateDiff("d", pdFecVencimiento, pdFecParaDeuda))
'    If vDiasAtra = 0 Then
'       vInteMora = 0
'       vCostCustMora = 0
'       vImpu = 0
'    Else
'       vInteMora = clsCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresVencido, vDiasAtra)
'       vInteMora = Round(vInteMora, 2)
'       vCostCustMora = clsCalculos.nCalculaCostoCustodiaMoratorio(pnValorTasacion, pnCostoCustodiaMoratoria, vDiasAtra)
'       vCostCustMora = Round(vCostCustMora, 2)
'       vImpu = (vInteMora + vCostCustMora) * pnTasaImpuesto
'       vImpu = Round(vImpu, 2)
'    End If
'    vDeud = Round(pnSaldoCapital, 2) + vInteMora + vCostCustMora + vImpu
'    If pnEstado = ColocEstado.gColPEstPRema Or pnEstado = ColocEstado.gColPEstAdjud Then      ' Si esta en via de Remate
'       vDeud = vDeud + Round((pnCostoPreparacionRemate * pnValorTasacion), 2)
'    End If
'    CalculaDeudaPrendario = vDeud
'    Set clsCalculos = Nothing
'End Function
'
'Public Function CalculaMinimoPagar(pnTasaInteres As Double, pnPlazo As Integer, pnValorTasacion As Double, _
'                                    pnTasaCustodia As Double, pnSaldoCapital As Double, pnEstado As ColocEstado, _
'                                    pdFecVencimiento As Date, pnTasaInteresVencido As Double, _
'                                    pnCostoCustodiaMoratoria As Double, pnTasaImpuesto As Double, _
'                                    pnCostoPreparacionRemate As Double, pgdFecSis As Date, Optional pdFecParaDeuda As String) As Double
'    ' **************************************
'    ' ** Calculo del Monto Minimo a Pagar **
'    ' **************************************
'    '    Dim NumeroDias As Single
'    Dim vFactor As Single
'    Dim vdiasAtraso As Double
'    Dim vInteres As Double
'    Dim vCostoCustodia As Double
'    Dim vImpuesto As Double
'    Dim vInteresVencido As Double
'    Dim vCostoCustodiaVencido As Double
'    Dim vCostoPreparacionRemate As Double
'    Dim vMontoMinimo As Double
'
'    Dim clsCalculos As NColPCalculos
'    Set clsCalculos = New NColPCalculos
'
'    vInteres = 0
'    vFactor = 1 - 1 / (1 + pnTasaInteres) ^ (pnPlazo / 360)
'    If Len(pdFecParaDeuda) <= 0 Then pdFecParaDeuda = pgdFecSis
'    vdiasAtraso = IIf(DateDiff("d", pdFecVencimiento, pdFecParaDeuda) <= 0, 0, DateDiff("d", pdFecVencimiento, pdFecParaDeuda))
'    vCostoCustodia = clsCalculos.nCalculaCostoCustodia(pnValorTasacion, pnTasaCustodia, pnPlazo)
'    vCostoCustodia = Round(vCostoCustodia, 2)
'    vInteres = pnSaldoCapital * vFactor
'    vInteres = Round(vInteres, 2)
'    If vdiasAtraso = 0 Then
'       vInteresVencido = 0
'       vCostoCustodiaVencido = 0
'    Else
'       vInteresVencido = clsCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresVencido, vdiasAtraso)
'       vInteresVencido = Round(vInteresVencido, 2)
'       vCostoCustodiaVencido = clsCalculos.nCalculaCostoCustodiaMoratorio(pnValorTasacion, pnCostoCustodiaMoratoria, vdiasAtraso)
'       vCostoCustodiaVencido = Round(vCostoCustodiaVencido, 2)
'    End If
'    vImpuesto = (vInteresVencido + vInteres + vCostoCustodia + pnCostoCustodiaMoratoria) * pnTasaImpuesto
'    vImpuesto = Round(vImpuesto, 2)
'
'    If pnEstado = ColocEstado.gColPEstPRema Or pnEstado = ColocEstado.gColPEstAdjud Then  ' Si esta en via de Remate
'       vCostoPreparacionRemate = Round((pnCostoPreparacionRemate * pnValorTasacion), 2)
'    End If
'
'    vMontoMinimo = vInteresVencido + vCostoCustodiaVencido + vInteres + vCostoCustodia + vImpuesto + vCostoPreparacionRemate
'    vMontoMinimo = Round(vMontoMinimo, 2)
'
'    CalculaMinimoPagar = vMontoMinimo
'    Set clsCalculos = Nothing
'End Function
'
'Public Function TieneMovDespuesApertura(ByVal sCuenta As String) As Boolean
'Dim oCapMov As DCapMovimientos
'Set oCapMov = New DCapMovimientos
'TieneMovDespuesApertura = oCapMov.TieneMovDespuesApertura(sCuenta)
'Set oCapMov = Nothing
'End Function


