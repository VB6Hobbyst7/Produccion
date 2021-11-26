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
'Option Explicit
Dim sCodUser As String
Dim CuentasPersonaPrd() As String
Dim CuentasPersonaCre() As String
Dim CuentasPersonaJud() As String

Public Function ImprimeBoletaCMAC(ByVal psTit As String, ByVal psText As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal nMoneda As Moneda, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional bPregImpSaldos As Boolean = True, Optional sCodage As String = "") As String

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

Dim lsmensaje As String * 39

ETIQ:

On Error GoTo Error

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
sIni = 0
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

' Declaracion para los mensajes
Dim clsGen As DGeneral
Set clsGen = New DGeneral
        
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACICA - AHORRO"; Space(19 - 2 + sSep + saux); "CMACICA - AHORRO"
        
If nMoneda = gMonedaNacional Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If
        
Print #nFicSal, ""
        
If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If
        
Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
        
lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
If lnCliAux = 0 Then
    If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(saux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(saux + sMax + sSep - Len(psCodCta)); "Cuenta:" & psCodCta
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
        
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    lnCliAux = 1
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(saux + sMax + sSep - Len(psCodCta)); "Cuenta:" & psCodCta
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
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(saux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
    Print #nFicSal, ""
End If

Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
If pbSaldoInt Then
    If bPregImpSaldos Then
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
        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
        If pbOpSaldoC Then
            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
        Else
            Print #nFicSal, ""
        End If
    End If
Else
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
End If

Print #nFicSal, ""


If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

lnTope = 3 - lnCliAux

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)

'lsMensaje = clsGen.GetMensajeBoletas("XXXXX" & gCapAhorros, sCodAge)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff
lnNumLinCmac = lnNumLinCmac + 1
 
'For saux = 1 To (lnTope - lnNumLinCmac)
'    Print #nFicSal, ""
'Next saux
Close nFicSal
Exit Function
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function



Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional bPregImpSaldos As Boolean = True, Optional ByVal psCodCMAC As String, _
        Optional sCodage As String = "", Optional pnMonedaCMAC As Integer = 1, Optional ByRef bImpreSaldos As Boolean = False, Optional psOpeCod As String = "", _
        Optional sUserBus As String = "", Optional bExtorno As Boolean = False, Optional ByVal sProxReno As String = "", Optional psDisponible As String = "", Optional psIntangible As String = "", _
        Optional psDispCta As String = "", Optional psIntaCta As String = "") As String

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

Dim lsmensaje As String * 39

'** Para la Impresion del Codigo Antiguo en las Boletas
Dim lsCodAntiguo As String
Dim loConstSistema As NConstSistemas
Set loConstSistema = New NConstSistemas
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As DGeneral
        Set loGeneral = New DGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
Set loConstSistema = Nothing

ETIQ:

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open sLpt For Output As nFicSal

'Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 0
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

'' Declaracion para los mensajes
Dim clsGen As DGeneral
Set clsGen = New DGeneral
        
        
'' liquidacion Trujillo
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACICA - AHORRO"; Space(18 + sSep + saux); "CMACICA - AHORRO"

If Len(Trim(psCodCta)) < 18 Then
    If pnMonedaCMAC = 1 Then
        Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
    Else
        Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
    End If
Else
    If Mid(psCodCta, 9, 1) = 1 Then
        Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
    Else
        Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
    End If
End If

If psNomAgeRem = "" Then
    Print #nFicSal, ""
Else
    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
End If
        
If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If
        
Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(saux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    
    lnCliAux = 1
End If
        
If pbSaldoInt Or pbCuenta Then
    If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
        Print #nFicSal, Tab(sIni); "Cuenta:" & lsCodAntiguo; Space(8 + sSep + saux); "Cuenta:" & lsCodAntiguo
        'lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & Space(1 + vEspacio) _
                                    & " Contrato  : " & lsCodAntiguo & Space(10) & Chr(10)
    Else
        Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + saux); "Cuenta:" & psCodCta
    End If
            'Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta
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
            Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(saux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto)); lsNegritaOff
            'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
            If psDisponible = "" And psIntangible = "" Then
                Print #nFicSal, ""
            End If
        End If


        If psDisponible <> "" And psIntangible <> "" Then
           'psDispCta psIntaCta
           'gPrnCondensadaOFF
           'gPrnCondensadaON
           Print #nFicSal, gPrnCondensadaON; Tab(sIni); ImpreCarEsp(Trim(psDispCta)); Space(saux + sSep + sMax - Len(Trim(psDispCta))); ImpreCarEsp(Trim(psDispCta))
           Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psIntaCta)); Space(saux + sSep + sMax - Len(Trim(psIntaCta))); ImpreCarEsp(Trim(psIntaCta))
           Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psDisponible)); Space(saux + sSep + sMax - Len(Trim(psDisponible))); ImpreCarEsp(Trim(psDisponible))
           Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psIntangible)); Space(saux + sSep + sMax - Len(Trim(psIntangible))); ImpreCarEsp(Trim(psIntangible)); gPrnCondensadaOFF
        End If



    If pbSaldoInt Then
        If bPregImpSaldos Then
            If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                bImpreSaldos = True
                Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
                If pbOpSaldoC Then
                    Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
                Else
                    Print #nFicSal, ""
                End If
            Else
                bImpreSaldos = False
                Print #nFicSal, ""
                Print #nFicSal, ""
                pbSaldoInt = False
            End If
        Else
            Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
            If pbOpSaldoC Then
                Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
            Else
                Print #nFicSal, ""
            End If
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

If Trim(sProxReno) <> "" Then
    
      Print #nFicSal, Tab(sIni); sProxReno; Space(sMax + 20 - Len(sProxReno)); sProxReno
End If


If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    If Trim(psDisponible) = "" And Trim(psIntangible) = "" Then
        Print #nFicSal, ""
    End If
End If

If Not psNumDias = "----" Then
    Print #nFicSal, Tab(sIni); "Nro Dias Interes"; Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)); psNumDias; Space(-1 + sSep); "Nro Dias Interes"; Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")); psNumDias
    lnTope = 4 - lnCliAux
Else
    lnTope = 3 - lnCliAux
End If


Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
If psCodUsuRem = "" Then
    If bExtorno Then
        Print #nFicSal, Tab(sIni); "USUARIO EXTORNA : " & ImpreCarEsp(sCodUser); Space(11 + sSep + saux); "USUARIO EXTORNA : " & ImpreCarEsp(sCodUser)
        Print #nFicSal, Tab(sIni); "USUARIO OPERADOR: " & ImpreCarEsp(sUserBus); Space(11 + sSep + saux); "USUARIO OPERADOR: " & ImpreCarEsp(sUserBus)
    Else
        Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)
    End If
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + saux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + saux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
End If

lsmensaje = clsGen.GetMensajeBoletas(psCodCta, sCodage)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff
lnNumLinCmac = lnNumLinCmac + 1

If Left(psOpeCod, 4) = "2605" Or Left(psOpeCod, 4) = "1270" Or Left(psOpeCod, 4) = "1070" Then
    Print #nFicSal, ""
ElseIf bExtorno = True Then
    For saux = 1 To (lnTope - lnNumLinCmac - 1)
        Print #nFicSal, ""
    Next saux
ElseIf sProxReno <> "" Then
    For saux = 1 To (lnTope - lnNumLinCmac - 1)
        Print #nFicSal, ""
    Next saux
ElseIf Trim(psDisponible) <> "" And Trim(psIntangible) <> "" Then
    Print #nFicSal, ""
Else
    For saux = 1 To (lnTope - lnNumLinCmac)
        Print #nFicSal, ""
    Next saux
End If


Close nFicSal
Exit Function
Error:
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
        Optional sLpt As String, Optional ByVal psCodCMAC As String)

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax  As Integer
Dim saux As Integer
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

Dim lsmensaje As String * 39

ETIQ:
'On Error GoTo Error

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
sIni = 0
sMax = 33
saux = 5


sFecha = Format$(dFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "###,##0.00")
sSDisp = Format$(pnSaldo, "###,##0.00")

lsMonInt = Format$(psInter, "###,##0.00")
LnMonTotal = Val(psInter) + Val(pnMonto)
lsMonTotal = Format(LnMonTotal, "###,##0.00")

Dim lsCadArg As String
Dim lsInteres As String


Dim clsGen As DGeneral

If psCodCMAC = "102" Then  '' liquidacion Lima
        'Print #nFicSal, Chr$(10);
        Print #nFicSal, lsNegritaOn; 'Activa Negrita
        'Print #nFicSal, Tab(sIni); ""
        Print #nFicSal, Tab(sIni); "CMACICA - AHORRO"; Space(19 + sSep + saux); "CMACICA - AHORRO"
        
        If Mid(psCodCta, 9, 1) = 1 Then
            Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(sMax + saux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(sMax + saux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
        Else
            Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt;
        End If
        
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
        
        Print #nFicSal, ""
        Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If sMax - Len(psCliente) < 2 Then psCliente = Mid(psCliente, 1, sMax - 2)
            Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(saux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
        Else
            
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax - 2)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax - 2)
            
            Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
            Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
            lnCliAux = 1
        End If
        
        Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + saux); "Cuenta:" & psCodCta
        
        psTit = Trim(psTit)
        psTit = CentrarCadena(psTit, 28)
        Print #nFicSal, lsNegritaOn; 'Activa Negrita
        Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); '"-----" & psTit & "-----"
        
        
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
            Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); 'NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
        Else
            Print #nFicSal, ""
        End If
        
        Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
        Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)
        
        
        Set clsGen = New DGeneral
        lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        
        Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); lsNegritaOff  ' Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
        lnCliAux = lnCliAux + 1
        
        'Print #nFicSal, Chr$(12)
        For saux = 1 To 4 - lnCliAux
            Print #nFicSal, ""
        Next saux
        
        Close nFicSal
        
        Exit Sub

Else  '' liquidacion trujillo
        
        Print #nFicSal, lsNegritaOn; 'Activa Negrita
        'Print #nFicSal, Tab(sIni); ""
        Print #nFicSal, Tab(sIni); "CMACICA - AHORRO"; Space(19 + sSep + saux); "CMACICA - AHORRO"
        
        If Mid(psCodCta, 9, 1) = 1 Then
            Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(sMax + saux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(sMax + saux - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
        Else
            Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(lsNroExt) - Len(Trim(sNomAge)) - Len("-DOLARES")) + lsNroExt;
        End If
        
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
        
        Print #nFicSal, ""
        Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If sMax - Len(psCliente) < 2 Then psCliente = Mid(psCliente, 1, sMax - 2)
            Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(saux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
        Else
            
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax - 2)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax - 2)
            
            Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
            Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
            lnCliAux = 1
        End If
        
        Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + saux); "Cuenta:" & psCodCta
        
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
    If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
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
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); 'NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)

Set clsGen = New DGeneral
lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing

Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); lsNegritaOff  ' Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
lnCliAux = lnCliAux + 1

'Print #nFicSal, Chr$(12)
For saux = 1 To 4 - lnCliAux
    Print #nFicSal, ""
Next saux

Close nFicSal

Exit Sub

End If


Error:
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
    GetInteres = Round(((nTasa / 36000 + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

Public Function GetTasaInteresEspecial(ByVal cPersCod As String, ByVal nEstado As Integer, ByVal nProd As Integer, ByVal nMoneda As Integer, ByVal dFecha As Date) As Recordset
Dim clsCap As DCapMovimientos, RSTEMP As Recordset
Set RSTEMP = New Recordset
Set clsCap = New DCapMovimientos
Set GetTasaInteresEspecial = New Recordset
Set RSTEMP = clsCap.GetDataInteresEspecial(cPersCod, nEstado, nProd, nMoneda, dFecha)
Set GetTasaInteresEspecial = RSTEMP
Set clsCap = Nothing

End Function

Public Function ValidaMontoApertura(ByVal nProducto As Producto, ByVal nPersoneria As PersPersoneria, _
        ByVal nMonto As Double, ByVal nMoneda As Moneda, Optional bOrdPag As Boolean = False) As String
Dim nValor As Double
Dim clsGen  As nCapDefinicion
Dim sDesc As String
Set clsGen = New nCapDefinicion
nValor = clsGen.GetMontoMinimoAperturaPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
Set clsGen = Nothing
If nValor > nMonto Then
    ValidaMontoApertura = "Monto de Apertura debe ser > ó = a " & Format$(nValor, "#,##0.00")
Else
    ValidaMontoApertura = ""
End If
End Function

Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsPers As Recordset, _
        ByVal sAgencia As String, ByVal nOperacion As CaptacOperacion, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional sNroDoc As String, Optional sCodIF As String, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional sPersLavDinero As String = "", Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, _
        Optional pbITFAplica As Boolean = True, Optional pnITFTipoExoneracion As Integer = 0, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional ByVal psITFOpeCod As String = "", Optional ByVal cAlias As String, Optional ByVal nFirmasMin As Integer, _
        Optional cPersTasaEspecial As String, Optional bEsTasaPermanente As Boolean, Optional SBENPERSLAVDINERO As String, _
        Optional nTpoPrograma As Integer = 0, Optional pnMovNro As Long = 0) As String

Dim clsMant As DCapMovimientos
Dim dlsMant As DCapMantenimiento
Dim nErrNum As Integer
Dim nErrDesc As String, sCuenta As String
Dim bTrans As Boolean
Dim nMovNro As Long
Dim lnMontoCV As Currency



bTrans = False
Set clsMant = New DCapMovimientos
Set dlsMant = New DCapMantenimiento
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True

If pbITFAplica And pnITFValor > 0 Then
    If (nOperacion = gAhoApeChq Or nOperacion = gPFApeChq) And pbITFAsumido = False Then
        sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo - pnITFValor, dFecha, nFirmas, nPersoneria, _
            nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, pnITFValor, nTpoPrograma)
            
    ElseIf (nOperacion = gAhoApeChq Or nOperacion = gPFApeChq) And pbITFAsumido = True Then
        sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
            nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, , nTpoPrograma)
            
    ElseIf Not (nOperacion = gAhoApeChq Or nOperacion = gPFApeChq) And pbITFAsumido = True Then
            sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
            nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, , nTpoPrograma)
            
    Else
        If psITFOpeCod = gITFCobroEfectivo Then
            sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
                nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, , nTpoPrograma)
        Else
            sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo - pnITFValor, dFecha, nFirmas, nPersoneria, _
                nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, , nTpoPrograma)
        End If
        
    End If
Else
    sCuenta = clsMant.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
        nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, cAlias, nFirmasMin, cPersTasaEspecial, bEsTasaPermanente, , nTpoPrograma)
End If

clsMant.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsMant.GetnMovNro(sMovNro)


Select Case nOperacion
    Case gAhoApeChq, gPFApeChq, gCTSApeChq
        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        
        Dim rstemp1 As ADODB.Recordset
         
        Set rstemp1 = clsMant.GetDatosCheque(sNroDoc, sCodIF, "01")
        
        clsMant.AgregaCuentaDocumento sCuenta, TpoDocCheque, sNroDoc, sCodIF, sMovNro, nMovNro, nSaldo, , rstemp1!cifcta
    
        If pbITFAplica And pnITFValor > 0 Then
            If pnITFTipoExoneracion = 0 Then
               If pbITFAsumido Then
                  clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, 0, nSaldo
                  clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFAsumido, pnITFValor
               Else
                  If psITFOpeCod = gITFCobroEfectivo Then
                     clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, 0, nSaldo
                     clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, pnITFValor
                  Else
                     clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, 0 - pnITFValor, nSaldo - pnITFValor
                     clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, pnITFValor
                  End If
               End If
            End If
        End If
    Case gAhoApeTransf, gPFApeTransf, gCTSApeTransf
        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        
        If pnMovNroTransfer <> -1 Then
            clsMant.AgregaMovRef nMovNro, pnMovNroTransfer
            If pnMonedaTrans <> nMoneda Then
                GetTipCambio dFecha
                
                If nMoneda = gMonedaNacional Then
                    clsMant.ActualizaMovPendientesRend pnMovNroTransfer, Round(nSaldo / gnTipCambioC, 2)
                    lnMontoCV = Round((nSaldo / gnTipCambioC) * gnTipCambio, 2)
                    clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
                    clsMant.AgregaMovTipoCambio nMovNro, gnTipCambioC
                    If gnTipCambioC < gnTipCambio Then
                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nSaldo)
                    Else
                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nSaldo)
                    End If
                Else
                    clsMant.ActualizaMovPendientesRend pnMovNroTransfer, Round(nSaldo * gnTipCambioV, 2)
                    lnMontoCV = Round(nSaldo * gnTipCambio, 2)
                    clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
                    clsMant.AgregaMovTipoCambio nMovNro, gnTipCambioV
                    If gnTipCambioV > gnTipCambio Then
                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nSaldo * gnTipCambioV, 2))
                    Else
                        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nSaldo * gnTipCambioV, 2))
                    End If
                End If
            Else
                clsMant.ActualizaMovPendientesRend pnMovNroTransfer, nSaldo
            End If
            
        End If
        'Aqui debe registrarse la referencia con la transferencia de caja general
    Case Else
        clsMant.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        clsMant.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo

        If pbITFAplica And pnITFValor > 0 Then
            If pnITFTipoExoneracion = 0 Then
               If pbITFAsumido Then
                  clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, nSaldo, nSaldo
                  clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFAsumido, pnITFValor
               Else
                  If psITFOpeCod = gITFCobroEfectivo Then
                     clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, nSaldo, nSaldo
                     clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, pnITFValor
                  Else
                     clsMant.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, nSaldo - pnITFValor, nSaldo - pnITFValor
                     clsMant.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, pnITFValor
                  End If
               End If
            End If
        End If
End Select

If pbITFAplica Then
    If pnITFTipoExoneracion <> 0 Then
        clsMant.AgregaCtaExoneradaITF sCuenta, pnITFTipoExoneracion, sMovNro
    End If
End If



If nOperacion = gCTSApeEfec Or nOperacion = gCTSApeChq Or nOperacion = gCTSApeTransf Then
    clsMant.AgregaMovParamCTS nMovNro, nPorcRetCTS
End If
''ALPA 20081010
''If sPersLavDinero <> "" Then
''    clsMant.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
''End If
pnMovNro = nMovNro
If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "233" Then
    Do While Not rsPers.EOF
        clsMant.AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4))), Left(rsPers("Firma Oblig"), 1)
        dlsMant.AgregaBitacoraPPer sMovNro, sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4))), Left(rsPers("Firma Oblig"), 1)
        rsPers.MoveNext
    Loop

Else
    Do While Not rsPers.EOF
        clsMant.AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4)))
        dlsMant.AgregaBitacoraPPer sMovNro, sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4))), "S"
        rsPers.MoveNext
    Loop
End If



clsMant.dbCmact.CommitTrans
bTrans = False


If Mid(sCuenta, 6, 3) = "233" Or Mid(sCuenta, 6, 3) = "234" Then
  Dim cAlcance As String
  Dim osorteo As DSorteo, rsDatos As ADODB.Recordset
  Dim gsMontoxCuponS As Double, gsMontoxCuponD As Double, gsMinPlazoOtorgado As Long, gsLimMaxOtorgadoS As Double, gsLimMaxOtorgadoD As Double, gsProcesaSorteo As Boolean
  
  Set osorteo = New DSorteo
  Set rsDatos = New ADODB.Recordset
   Set rsDatos = osorteo.GetParametrosSorteo
      While Not rsDatos.EOF
            
            Select Case rsDatos!nparamcod
               Case 2001
                     gsMontoxCuponS = CDbl(rsDatos!cparamvalor)
               
               Case 2002
                    gsMontoxCuponD = CDbl(rsDatos!cparamvalor)
                    
               Case 2003
                    gsMinPlazoOtorgado = CLng(rsDatos!cparamvalor)
                    
               Case 2004
                    gsLimMaxOtorgadoS = CDbl(rsDatos!cparamvalor)
               
               Case 2005
                    gsLimMaxOtorgadoD = CDbl(rsDatos!cparamvalor)
                    
                Case 2008
                    gsProcesaSorteo = IIf(rsDatos!cparamvalor = 0, False, True)
                         
            
            End Select
            rsDatos.MoveNext
      Wend
  
      If gsProcesaSorteo = True Then
        cAlcance = ""
        If Mid(sCuenta, 4, 2) = "13" Or Mid(sCuenta, 4, 2) = "03" Or Mid(sCuenta, 4, 2) = "06" Then
          cAlcance = "13"
        ElseIf Mid(sCuenta, 4, 2) = "04" Or Mid(sCuenta, 4, 2) = "05" Then
          cAlcance = "04"
        ElseIf Mid(sCuenta, 4, 2) = "01" Or Mid(sCuenta, 4, 2) = "14" Then
          cAlcance = "01"
        ElseIf Mid(sCuenta, 4, 2) <> "13" And Mid(sCuenta, 4, 2) <> "03" And Mid(sCuenta, 4, 2) <> "06" And Mid(sCuenta, 4, 2) <> "04" And Mid(sCuenta, 4, 2) <> "05" And Mid(sCuenta, 4, 2) <> "01" And Mid(sCuenta, 4, 2) <> "14" Then
          cAlcance = Mid(sCuenta, 4, 2)
        End If
        
         Set rsDatos = New ADODB.Recordset
         Set rsDatos = osorteo.GetSorteo("I", cAlcance)
          
         If Not rsDatos.EOF Then
            If CInt(Right(rsDatos!CnumSorteo, 2)) = 1 Then
                 Call osorteo.ProcesarCtaSorteoPrimer(rsDatos!CnumSorteo, gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, cAlcance, sMovNro, sCuenta)
            Else
                 Call osorteo.ProcesarCtaSorteoOtros(rsDatos!CnumSorteo, gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, cAlcance, sMovNro, sCuenta)
            End If
         End If
                  
        If Mid(sCuenta, 4, 2) = "01" Or Mid(sCuenta, 4, 2) = "14" Then
                  
                Set rsDatos = New ADODB.Recordset
                Set rsDatos = osorteo.GetSorteo("I", "00")
                     
                 If Not rsDatos.EOF Then
                   If CInt(Right(rsDatos!CnumSorteo, 2)) = 1 Then
                     Call osorteo.ProcesarCtaSorteoPrimer(rsDatos!CnumSorteo, gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, "00", sMovNro, sCuenta)
                   Else
                     Call osorteo.ProcesarCtaSorteoOtros(rsDatos!CnumSorteo, gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, "00", sMovNro, sCuenta)
                   End If
                 End If
        End If
                 
          
     End If

End If








CapAperturaCuenta = sCuenta


Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    CapAperturaCuenta = ""
    Err.Raise Err.Number, "CapAperturaCuenta", Err.Description
End Function

Public Sub CapAperturaCuentaLote(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsCuenta As Recordset, _
    ByVal nOperacion As CaptacOperacion, ByVal sGlosa As String, ByVal sMovNro As String, ByVal nTipoTasa As CaptacTipoTasa, _
    Optional bCheque As Boolean = False, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
    Optional dFechaValorizacion As Date, Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
    Optional sNomCmac As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", _
    Optional pbITFAplica As Boolean = True, Optional pnITFTipoExoneracion As Integer = 0, Optional pbITFAsumido As Boolean = False, Optional ByVal psITFOpeCod As String = "")

Dim clsMov As DCapMovimientos
Dim clsMant As NCapMantenimiento
Dim nErrNum As Integer
Dim nErrDesc As String, sCuenta As String, sCtaCodAbono As String
Dim bTrans As Boolean, bOrdPag As Boolean, bCtaAboInt As Boolean
Dim nMovNro As Long, nPlazo As Long, nFirmas As Long
Dim nSaldo As Double, nTasa As Double
Dim nitf As Double

Dim nFormaRetiro As CaptacPFFormaRetiro
Dim nPersoneria As PersPersoneria
Dim nTipoCuenta As ProductoCuentaTipo
Dim sImpresion As String, sCodUser As String, sAgencia As String, sOperacion As String
Dim dFecSis As Date
Dim sCad As String, sNomTit As String, sPersLavDinero As String, sBoletaLavDinero As String
Dim clsImp As NCapImpBoleta
Dim rsProdPers As Recordset
Dim lbImpRegFirma As Byte
Dim oGen As DGeneral
Dim nMontoLavDinero As Double, nTC As Double
Dim clsLav As nCapDefinicion
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim rsPers As Recordset

'Realiza la Validación para el Lavado de Dinero
sPersLavDinero = ""
sBoletaLavDinero = ""

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
sAgencia = Mid(sMovNro, 18, 2)

Set clsLav = New nCapDefinicion
nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
Set clsLav = Nothing

If nMoneda = gMonedaNacional Then
    Dim clsTC As nTipoCambio
    Set clsTC = New nTipoCambio
    nTC = clsTC.EmiteTipoCambio(dFecSis, TCFijoDia)
    Set clsTC = Nothing
Else
    nTC = 1
End If

Set oGen = New DGeneral

lbImpRegFirma = CInt(oGen.LeeConstSistema(100))
Set oGen = Nothing

bTrans = False
Set clsMov = New DCapMovimientos
On Error GoTo ErrGraba


clsMov.dbCmact.BeginTrans 'Inicializa la transacción
bTrans = True
    
'Agrega el movimiento
clsMov.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsMov.GetnMovNro(sMovNro)
If nOperacion = gCTSApeLoteEfec Or nOperacion = gCTSApeLoteChq Then
    clsMov.AgregaMovParamCTS nMovNro, nPorcRetCTS
End If
clsMov.dbCmact.CommitTrans 'Inicializa la transacción
bTrans = False


'Variables globales para las aperturas en lote
bCtaAboInt = False
nTipoCuenta = gPrdCtaTpoIndiv
sCtaCodAbono = ""
nFirmas = 1

Set clsMant = New NCapMantenimiento
If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    clsMov.AgregaMovDoc nMovNro, TpoDocCheque, sNroDoc, dFecSis
    sOperacion = "APERTURA LOTE CHEQUE"
Else
    sOperacion = "APERTURA LOTE EFECTIVO"
End If
Do While Not rsCuenta.EOF
    clsMov.dbCmact.BeginTrans
    bTrans = True
    nSaldo = CDbl(rsCuenta("Monto"))
    If IsNumeric(rsCuenta("ITF")) Then
        nitf = CDbl(rsCuenta("ITF"))
    Else
        nitf = 0
    End If
    nPersoneria = CInt(rsCuenta("cPersoneria"))
    nTasa = CDbl(rsCuenta("nTasaNominal"))
    
    nPlazo = 0
    bOrdPag = False
    If nProducto = gCapPlazoFijo Then
        nPlazo = rsCuenta("Plazo")
        nFormaRetiro = CLng(Right(rsCuenta("Forma Retiro"), 4))
    ElseIf nProducto = gCapAhorros Then
        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
    End If
    
    If pbITFAplica And nitf > 0 Then
        If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or pbITFAsumido Then
            sCuenta = clsMov.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecSis, nFirmas, nPersoneria, _
                nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
        Else
            If psITFOpeCod = gITFCobroEfectivo Then
                sCuenta = clsMov.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecSis, nFirmas, nPersoneria, _
                    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
            Else
                sCuenta = clsMov.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo - nitf, dFecSis, nFirmas, nPersoneria, _
                    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
            End If
        End If
    Else
        sCuenta = clsMov.AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecSis, nFirmas, nPersoneria, _
            nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
    End If
    
    rsCuenta("Cuenta") = sCuenta
    rsCuenta.Update
    
    Select Case nOperacion
        Case gAhoApeLoteChq, gPFApeLoteChq, gCTSApeLoteChq
            clsMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
            clsMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
            clsMov.AgregaDocRecCapta sCuenta, TpoDocCheque, sNroDoc, sCodIF, nMovNro, nSaldo
        
            If pbITFAplica And nitf > 0 Then
                If pnITFTipoExoneracion = 0 Then
                   If pbITFAsumido Then
                      clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo, nSaldo
                      clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFAsumido, nitf
                   Else
                      If psITFOpeCod = gITFCobroEfectivo Then
                         clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo, nSaldo
                         clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, nitf
                      Else
                         clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo - nitf, nSaldo - nitf
                         clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, nitf
                      End If
                   End If
                End If
            End If
        
        Case Else
            clsMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
            clsMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
    
            If pbITFAplica And nitf > 0 Then
                If pnITFTipoExoneracion = 0 Then
                   If pbITFAsumido Then
                      clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo, nSaldo
                      clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFAsumido, nitf
                   Else
                      If psITFOpeCod = gITFCobroEfectivo Then
                         clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo, nSaldo
                         clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, nitf
                      Else
                         clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, nitf, nSaldo - nitf, nSaldo - nitf
                         clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, nitf
                      End If
                   End If
                End If
            End If
    End Select
    
    If pbITFAplica Then
        If pnITFTipoExoneracion <> 0 Then
            clsMov.AgregaCtaExoneradaITF sCuenta, pnITFTipoExoneracion, sMovNro
        End If
    End If
    
    clsMov.AgregaNuevoProdPers sCuenta, rsCuenta("Código"), gCapRelPersTitular
    
    If nSaldo >= Round(nMontoLavDinero * nTC, 2) Then
        sPersLavDinero = rsCuenta("Código")
        clsMov.AgregaMovLavDinero nMovNro, sPersLavDinero
        
        Set rsPers = clsMant.GetDatosPersona(sPersLavDinero)
        sDireccion = rsPers("Direccion")
        sDocId = rsPers("ID N°")
        sNombre = rsCuenta("Nombre")
        rsPers.Close
        Set rsPers = Nothing
        Set clsImp = New NCapImpBoleta
        If sBoletaLavDinero <> "" Then sBoletaLavDinero = sBoletaLavDinero & Chr$(12)
        sBoletaLavDinero = sBoletaLavDinero & clsImp.ImprimeBoletaLavadoDinero(sNomCmac, sNomAge, dFecSis, sCuenta, sNombre, sDocId, sDireccion, _
                    sNombre, sDocId, sDireccion, sNombre, sDocId, sDireccion, sOperacion, nSaldo, sLpt, , True, "INDIVIDUAL")
        Set clsImp = Nothing
    End If
    
    
    clsMov.dbCmact.CommitTrans
    bTrans = False
    rsCuenta.MoveNext
Loop

Set clsMov = Nothing
sCad = ""

'Imprime registro de firmas
rsCuenta.MoveFirst
Do While Not rsCuenta.EOF
    bOrdPag = False
    If nProducto = gCapAhorros Then
        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
    End If
    If sCad <> "" Then sCad = sCad & Chr$(12)
    'Creamos el recordset para el registro de firmas
    Set rsProdPers = New Recordset
    rsProdPers.Fields.Append "Codigo", adVarChar, 13
    rsProdPers.Fields.Append "Relacion", adVarChar, 30
    rsProdPers.Open
    rsProdPers.AddNew
    rsProdPers("Codigo") = rsCuenta("Código")
    rsProdPers("Relacion") = "TITULAR"
    rsProdPers.Update
    If lbImpRegFirma = 1 Then
        sCad = sCad & clsMant.GeneraRegistroFirmas(rsCuenta("Cuenta"), "INDIVIDUAL", dFecSis, bOrdPag, rsProdPers, sNomAge, dFecSis)
    End If
    Set rsProdPers = Nothing
    rsCuenta.MoveNext
Loop

Set clsImp = New NCapImpBoleta


'Imprime Certificado de plazos fijos
If nProducto = gCapPlazoFijo Then
    Dim sDirCli As String
    Dim sFormaRetiro As String
    Dim sLetras As String
    
    Dim clsBol As NCapImpBoleta
    rsCuenta.MoveFirst
    Do While Not rsCuenta.EOF
        nSaldo = CDbl(rsCuenta("Monto"))
        nTasa = CDbl(rsCuenta("nTasaNominal"))
        nPlazo = rsCuenta("Plazo")
        nFormaRetiro = CLng(Right(rsCuenta("Forma Retiro"), 4))
        sNomTit = clsMant.GetNombreTitulares(rsCuenta("Cuenta"))
        Set rsPers = clsMant.GetDatosPersona(rsCuenta("Código"))
        sDirCli = Trim(rsPers("Direccion")) & ""
        sLetras = ConversNL(nMoneda, nSaldo)
        sFormaRetiro = Trim(Left(rsCuenta("Forma Retiro"), 25))
        'If scad <> "" Then scad = scad & Chr$(12)
        sCad = sCad & clsImp.ImprimeCertificadoPlazoFijo(dFecSis, sNomTit, sDirCli, rsCuenta("Cuenta"), "1", nPlazo, nSaldo, _
                    nTasa, sFormaRetiro, sLetras, , True)
        rsCuenta.MoveNext
    Loop
End If


Dim clsPrev As previo.clsprevio
Set clsPrev = New previo.clsprevio
clsPrev.Show sCad, "APERTURA EN LOTE", False
    
    
'Imprime las boletas
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
sTipDep = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES")
If bCheque Then
    sModDep = "Depósito Cheque"
Else
    sModDep = "Depósito Efectivo"
End If
Select Case nProducto
    Case gCapAhorros
        sTipApe = "APERTURA AHORROS"
    Case gCapPlazoFijo
        sTipApe = "APERTURA PLAZO FIJO"
    Case gCapCTS
        sTipApe = "APERTURA CTS"
End Select

Do

sCad = Chr$(27) & Chr$(77)
rsCuenta.MoveFirst
Dim lnSaldo As Double, lnsaldisp As Double

Do While Not rsCuenta.EOF
    bOrdPag = False
    If nProducto = gCapAhorros Then
        bOrdPag = IIf(rsCuenta("OP") = "1", True, False)
    End If
    If bOrdPag Then sTipApe = "APERTURA AHORROS CON OP"
    sNomTit = clsMant.GetNombreTitulares(rsCuenta("Cuenta"))
    nSaldo = CDbl(rsCuenta("Monto"))
    If IsNumeric(rsCuenta("ITF")) Then
        nitf = CDbl(rsCuenta("ITF"))
    Else
        nitf = 0
    End If
    
    'If Right(sCad, 1) <> Chr$(12) And sCad <> "" Then sCad = sCad & Chr$(12)
    If bCheque Then
        sCad = sCad & ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, Trim(nOperacion), Trim(nSaldo), sNomTit, rsCuenta("Cuenta"), Format$(dFechaValorizacion, "dd/mm/yyyy"), nSaldo, 0, "Fecha Valor", 1, nSaldo, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, False, , , , True)
        lnsaldisp = 0
        lnSaldo = nSaldo
    Else
        sCad = sCad & ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), Trim(sCodOpe), Trim(nSaldo), sNomTit, rsCuenta("Cuenta"), "", nSaldo, 0, "", 1, nSaldo, , , , , , , , , , dFecSis, Trim(sNomAge), sCodUser, sLpt, False, , , , True)
        lnsaldisp = nSaldo
        lnSaldo = nSaldo
    End If
    
    If pbITFAplica And nitf > 0 Then
       ' sCad = sCad & Chr$(12)
        If pbITFAsumido Then
            If psITFOpeCod = gITFCobroEfectivo Then
'                sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA",  False, sNomAge, sCodUser)
                 'sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser) & Chr$(12)
                 fgITFImprimeBoleta sNomTit, nitf, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , , Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, True, lnsaldisp, lnSaldo, psITFOpeCod
            Else
'                sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser)
                 'sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser) & Chr$(12)
                 fgITFImprimeBoleta sNomTit, nitf, sGlosa, dFecSis, sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, True, lnsaldisp, lnSaldo, CStr(nOperacion)
            End If
        Else
            If psITFOpeCod = gITFCobroEfectivo Then
'                sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser)
                 'sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser) & Chr$(12)
                 fgITFImprimeBoleta sNomTit, nitf, sGlosa, dFecSis, sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR EL CLIENTE", False, sNomAge, sCodUser, True, lnsaldisp, lnSaldo, CStr(nOperacion)
            Else
'                sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser)
               '  sCad = sCad & fgITFImprimeBoletaCadena(sNomTit, nITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser) & Chr$(12)
               fgITFImprimeBoleta sNomTit, nitf, sGlosa, dFecSis, sLpt, 1, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR EL CLIENTE", False, sNomAge, sCodUser, True, lnsaldisp, lnSaldo, CStr(nOperacion)
            End If
        End If
    End If
    
    rsCuenta.MoveNext
Loop

Loop Until MsgBox("Desea reimprimir boletas", vbYesNo + vbQuestion, "AVISO") = vbNo
'If sBoletaLavDinero <> "" Then sCad = sCad & Chr$(12) & sBoletaLavDinero
'clsPrev.Show sCad, "APERTURA EN LOTE", False

Set clsImp = Nothing
Set clsMant = Nothing
Exit Sub

ErrGraba:
    If bTrans Then clsMov.dbCmact.RollbackTrans
    Set clsMov = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function ValidaCuentaOperacion(ByVal sCuenta As String, Optional bAbono As Boolean = False, _
            Optional bGarantia As Boolean = False) As String
Dim clsCap As DCapMovimientos
Dim nEstado As CaptacEstado
Dim rsBlq As Recordset
Dim sCad As String
Dim bPrimero As Boolean
Set clsCap = New DCapMovimientos
nEstado = clsCap.GetEstadoCuenta(sCuenta)
If nEstado < 0 Then
    ValidaCuentaOperacion = "Cuenta No Existe"
Else
    Select Case nEstado
    
        Case gCapEstAnulada, gCapEstCancelada
            ValidaCuentaOperacion = "Cuenta CANCELADA O ANULADA"
        Case gCapEstBloqTotal, gCapEstBloqRetiro
            sCad = ""
            Set rsBlq = clsCap.GetCuentaBloqueos(sCuenta, gCapTpoBlqRetiro)
            If Not bAbono Then
                If Not (rsBlq.EOF And rsBlq.BOF) Then
                    bPrimero = True
                    Do While Not rsBlq.EOF
                        If Not (rsBlq("nBlqMotivo") = gCapMotBlqRetGarantia And bGarantia) Then
                            If bPrimero Then
                                sCad = "BLOQUEO PARA RETIRO: " & Chr$(13)
                                bPrimero = False
                            End If
                            sCad = sCad & "- " & Trim(rsBlq("Motivo")) & Chr$(13)
                        End If
                        rsBlq.MoveNext
                    Loop
                End If
            End If
            Set rsBlq = clsCap.GetCuentaBloqueos(sCuenta, gCapTpoBlqTotal)
            If Not (rsBlq.EOF And rsBlq.BOF) Then
                bPrimero = True
                Do While Not rsBlq.EOF
                    If Not (rsBlq("nBlqMotivo") = gCapMotBlqTotGarantia And bGarantia) Then
                        If bPrimero = True Then
                            sCad = sCad & "BLOQUEO TOTAL: " & Chr$(13)
                            bPrimero = False
                        End If
                        sCad = sCad & "- " & Trim(rsBlq("Motivo")) & Chr$(13)
                    End If
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
Dim dAuxiliar As Date, dRenovacion As Date, dVencimiento As Date
Dim nPlazo As Long, nFormaRetiro As Integer
Dim rsCta As Recordset
Set clsCap = New DCapMantenimiento
Set rsCta = clsCap.GetDatosCuentaPF(sCuenta)
Set clsCap = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    dAuxiliar = rsCta("dAuxiliar")
    dRenovacion = rsCta("dRenovacion")
    nPlazo = rsCta("nPlazo")
    dVencimiento = DateAdd("d", nPlazo, dRenovacion)
    nFormaRetiro = rsCta("nformaretiro")
    If DateDiff("d", dVencimiento, dFecSis) <> 0 Then
    
        'If DateDiff("d", dRenovacion, dFecSis) < 30 Then
        '    sCad = "El número de dias transcurridos desde la APERTURA/RENOVACION es MENOR A 30 DIAS"
        'Else
           
            If DateDiff("d", dAuxiliar, dFecSis) < 30 And nFormaRetiro = gCapPFFormRetMensual Then
                sCad = "El número de dias transcurridos desde la última " & Chr$(13) _
                        & "FECHA PACTADA DE RETIRO es MENOR A 30 DIAS"
            ElseIf DateDiff("d", dAuxiliar, dFecSis) < 30 And nFormaRetiro = 4 Then
                sCad = "El número de dias transcurridos desde la última " & Chr$(13) _
                        & "FECHA PACTADA DE RETIRO es MENOR A 30 DIAS"
                        
            ElseIf DateDiff("d", dRenovacion, dFecSis) < nPlazo And nFormaRetiro = gCapPFFormRetFinalPlazo Then
                sCad = "El número de dias transcurridos desde la última " & Chr$(13) _
                        & " RENOVACION es MENOR A 30 DIAS"
            Else
                sCad = ""
            End If
        'End If
    End If
Else
    sCad = "Cuenta NO Existe"
End If
rsCta.Close
Set rsCta = Nothing
ValidaRetiroIntPF = sCad
End Function

Public Function CapYaRetiro(ByVal psCuenta As String, ByVal dFecSis As Date) As Boolean
 Dim FRetiro1 As Date
    
    FRetiro1 = GetFechaUltimoRetiroIntPF(psCuenta)

    
    
    
    

End Function

Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
        Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
        Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
        Optional sCodCMAC As String = "", Optional sNomCmac As String = "", Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
        Optional bImprimeBoleta As Boolean = True, Optional pnMovNroTransfer As Long = -1, _
        Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional ByVal psCodCMAC As String, Optional pCon As ADODB.Connection = Nothing, _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional sOPedesc As String = "", Optional ByRef pnExtracto As Long, _
        Optional SBENPERSLAVDINERO As String) As Double

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
Dim sMsgOpe As String, sCodUser As String, sAgencia As String
Dim nMovNro As Long
Dim dFecSis As Date
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas
Dim lnSaldoDisp As Double
Dim lnSaldoCont As Double
Dim bImpreSaldos As Boolean


Dim lnMontoCV As Currency

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
sAgencia = Mid(sMovNro, 18, 2)

Set clsCap = New DCapMovimientos

Dim MontoItf As Double
nMovNro = clsCap.GetnMovNro(sMovNro)

'MontoItf = 0
'If pbITFAplica And Mid(sCuenta, 6, 3) <> CStr(gCapCTS) And pbITFAsumido = False And psITFOperacion = CStr(gITFCobroCargo) Then     'Filtra para CTS
'        If nMonto > gnITFMontoMin Then
'            If psITFOperacion = CaptacOperacion.gITFCobroCargo Then
'                MontoItf = clsCap.GetITFMOV(nMovNro)
'            End If
'        End If
'         nSaldoCnt = nSaldoCnt - MontoItf
'         nSaldoDisp = nSaldoDisp - MontoItf
'End If


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
bInamovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
pnExtracto = nExtracto

rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos

On Error GoTo ErrGraba

If Not pCon Is Nothing Then
    clsCap.SetConexion pCon
    clsCap.bTransaccion = True
    bTransaccion = True
End If

If Not bTransaccion Then clsCap.dbCmact.BeginTrans
bTrans = True


If Not ValidaEstadoCuenta(sCuenta, False) Then   'Valida el estado de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia.", vbInformation, "Aviso"
    CapAbonoCuentaAho = 0
    Exit Function
End If

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
    If nTipoDoc = TpoDocCheque And nOperacion <> "200252" Then 'Si el abono es con cheque
        'If pnITFOperacion = CaptacOperacion.gITFCobroCargo Then
            nSaldoDisp = nSaldoDisp + 0
            nIntGanado = 0
        'End If
        clsCap.ActualizadUltContactoAho sCuenta, dFecSis
    ElseIf nTipoDoc = TpoDocCheque And nOperacion = "200252" Then
    
        clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        clsCap.ActualizadUltContactoAho sCuenta, dFecSis
        
    ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
        clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
    End If
End If

'******** para Boleta ITF

 lnSaldoDisp = nSaldoDisp
 lnSaldoCont = nSaldoCnt
'-----



clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        If psItfOperacion = gITFCobroEfectivo Then
            clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        Else
            clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
            clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        End If
    End If
End If

If sNroDoc = "" Then ' Si la operacion es en efectivo
    If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
    Else
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
    End If
    
    If nOperacion = gAhoDepEfec Then
        sMsgOpe = "Depósito Efectivo"
    ElseIf nOperacion = gAhoDepTransf Then
        sMsgOpe = "Depósito Transferencia"
    ElseIf nOperacion = gAhoDepDevFondoGar Then
        sMsgOpe = "Depósito Dev Fondo Gar"
    ElseIf nOperacion = gAhoDepPagServEdelnor Then
        sMsgOpe = "Depósito Serv Edelnor"
    ElseIf nOperacion = gAhoDepPagServHidrandina Then
        sMsgOpe = "Depósito Serv Hidrandina"
    ElseIf nOperacion = gAhoDepPagServSedalib Then
        sMsgOpe = "Depósito Serv Sedalib"
    ElseIf nOperacion = gAhoDepSobCaja Then
        sMsgOpe = "Depósito Sobrante"
    ElseIf nOperacion = gAhoDepAboOtrosConceptos Then
        sMsgOpe = "Depósito Otros Conceptos"
    ElseIf nOperacion = gAhoDepDevCredPersonales Then
        sMsgOpe = "Depósito Dev Cred Pers"
    ElseIf nOperacion >= 200247 And nOperacion <= 200251 Then
        sMsgOpe = Trim(sOPedesc)
    
    End If
Else
    If nTipoDoc = TpoDocCheque Then  'Si el abono es con cheque
    
       If nOperacion <> "200252" Then
          If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
               clsCap.ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, (pnITFValor * -1), nIntGanado, dUltMov, sMovNro, True
          Else
               clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
          End If
       Else
          If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
                clsCap.ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
          Else
                clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
          End If
       End If
         
       If nOperacion <> "200252" Then
             Dim rstemp1 As ADODB.Recordset
             
             Set rstemp1 = clsCap.GetDatosCheque(sNroDoc, sCodIF, "01")
            
             clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto, , rstemp1!cifcta
       End If
        
         sMsgOpe = "Dep.Chq"
    End If
End If

If bInamovilizada Then
    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
        clsCap.AgregaMovCap nMovNro, gAhoEstInmovAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInmovilizadaAct sCuenta
        clsCap.ActualizaInactivaAct sCuenta
    End If
End If

If bInactiva And Not bInamovilizada Then
    clsCap.AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    clsCap.ActualizaInactivaAct sCuenta
End If

If pnMovNroTransfer <> -1 Then
    clsCap.AgregaMovRef nMovNro, pnMovNroTransfer
    If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis
        
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
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
End If

CapAbonoCuentaAho = nSaldoCnt

If (Not bImprimeBoleta) Or nOperacion = gAhoDepEntConv Then
        Exit Function
End If

'****************************************
'*************************************************
If Not bTransaccion Then clsCap.dbCmact.CommitTrans
bTrans = False


Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "DEPOSITO AHORROS"

'If nOperacion >= 200247 And nOperacion <= 200251 Then
'        sTipApe = UCase(Trim(sOPedesc))
'End If

Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

If pbITFAplica And pnITFValor > 0 And pbITFAsumido = False Then
    
        If (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
            nSaldoDisp = nSaldoDisp - pnITFValor
            nSaldoCnt = nSaldoCnt - pnITFValor
            
        End If
End If

Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocCheque
                If sCodCMAC <> "" Then
                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion)
                Else
                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion)
                End If
            Case TpoDocNotaAbono
                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion)
        End Select
    Else
        If sCodCMAC <> "" Then
            sTipApe = "DEPOSITO CMAC AHORROS"
            ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion)
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion)
        End If
    End If
        
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            End If
        Else
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            End If
        End If
    End If
    
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo


Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function

Public Function CapAbonoCuentaCTS(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nPorcDisp As Double = 0, Optional nTipoDoc As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional pnMovNroTransfer As Long = -1, _
            Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional ByVal psCodCMAC As String, _
            Optional SBENPERSLAVDINERO As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nSaldoRetiro As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bTrans As Boolean, bInactiva As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim lnMontoCV As Currency

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

If Not ValidaEstadoCuenta(sCuenta, False) Then   'Valida el estado de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia.", vbInformation, "Aviso"
    CapAbonoCuentaCTS = 0
    Exit Function
End If

'Calcula Intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
'nIntSaldo = GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
nSaldoCnt = nSaldoCnt + nMonto
dUltMov = DateAdd("d", -1, dFecSis)

If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)

If sNroDoc = "" Then 'Si es abono en efectivo
    nSaldoDisp = nSaldoDisp + nMonto
    nSaldoRetiro = nSaldoRetiro + (nMonto * nPorcDisp / 100)
Else
    If nTipoDoc = TpoDocCheque Then 'Si es con cheque
        nSaldoDisp = nSaldoDisp + 0
    End If
End If

clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvision, nIntGanado
clsCap.AgregaMovParamCTS nMovNro, nPorcDisp
If sNroDoc = "" Then 'Si es abono en efectivo
    clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo
    clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    
    If nOperacion = gCTSDepEfec Then
        sMsgOpe = "Depósito Efectivo"
    ElseIf nOperacion = gCTSDepTransf Then
        sMsgOpe = "Depósito Transferencia"
    End If
Else
    If nTipoDoc = TpoDocCheque Then 'Si es con cheque
        clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        
         Dim rstemp1 As ADODB.Recordset
         
        Set rstemp1 = clsCap.GetDatosCheque(sNroDoc, sCodIF, "01")
       
        clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto, , rstemp1!ctaif
        
        sMsgOpe = "Depósito Cheque"
    End If
End If


If pnMovNroTransfer <> -1 Then
    clsCap.AgregaMovRef nMovNro, pnMovNroTransfer
    If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis
        
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
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
End If

'***************************
'*****************************************
clsCap.dbCmact.CommitTrans
bTrans = False
CapAbonoCuentaCTS = nSaldoCnt

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

Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocCheque
                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC
        End Select
    Else
        If sCodCMAC <> "" Then
            sTipApe = "DEPOSITO CMAC CTS"
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo


Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    
End Function

Public Function CapCargoAhoOPCertifcada(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal sNroDoc As String, Optional sCodIF As String = "", _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional SBENPERSLAVDINERO As String = "") As Double

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
clsCap.AgregaOrdenPago sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
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
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
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
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoAhoOPCertifcada = 0
End Function

Public Function CapCargoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
            Optional psCtaBanco As String = "", Optional pnMoneda As Moneda = gMonedaNacional, _
            Optional ByVal psCodCMAC As String, Optional pCon As ADODB.Connection = Nothing, _
            Optional ByVal psCodAge As String, Optional ByVal pbProcesaAutorizacion As Boolean = False, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional ByVal sOPedesc As String = "", Optional ByVal SBENPERSLAVDINERO As String = "") As Double

Dim lnOperacion As CaptacOperacion
Dim lnMonto As Double

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
Dim lnSaldoCont As Double
Dim lnSaldoDisp As Double
Dim bImpreSaldos As Boolean


Dim lsCodCliente As String

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False

Set clsCap = New DCapMovimientos


'MontoItf = 0
'If pbITFAplica And Mid(sCuenta, 6, 3) <> CStr(gCapCTS) And pbITFAsumido = False And psITFOperacion = CStr(gITFCobroCargo) Then     'Filtra para CTS
'        If nMonto > gnITFMontoMin Then
'            If psITFCargo = True Then
'                MontoItf = clsCap.GetITFMOV(nMovNro)
'            End If
'        End If
'         nSaldoCnt = nSaldoCnt - MontoItf
'         nSaldoDisp = nSaldoDisp - MontoItf
'End If
'
'

'Dim MontoItf As Double
'nMovNro = clsCap.GetnMovNro(sMovNro)


Dim lsrt As ADODB.Recordset
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set lsrt = clsMant.GetProductoPersona(sCuenta)
Set clsMant = Nothing

lsCodCliente = lsrt("cPersCod")

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
bInmovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos


On Error GoTo ErrGraba

If Not pCon Is Nothing Then
    clsCap.SetConexion pCon
    clsCap.bTransaccion = True
    bTransaccion = True
End If


If Not bTransaccion Then clsCap.dbCmact.BeginTrans
bTrans = True

Randomize

For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    CapCargoCuentaAho = 0
    Exit Function
End If

If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If


If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    CapCargoCuentaAho = 0
    Exit Function
End If

'----------------------------------------

'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)

clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto

lnSaldoDisp = nSaldoDisp
lnSaldoCont = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        If psItfOperacion = gITFCobroEfectivo Then
            clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
            clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        Else
            clsCap.ActualizaCargoCaptacion sCuenta, nMonto + pnITFValor, nMonto + pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
            clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
            clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        End If
    End If
Else
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
End If

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        Dim nEstadoOP As CaptacOrdPagoEstado
        nEstadoOP = gCapOPEstCobrada
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certif. OP"
            nEstadoOP = gCapOPEstCertifiCada
        End If
        
        clsCap.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, Right(sCodIF, 13), sMovNro, nMovNro, nMonto, nEstadoOP
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
    ElseIf nOperacion = gAhoRetComServHidrandina Then
        sMsgOpe = "Retiro Serv Hidrandina"
    ElseIf nOperacion = gAhoRetComServSEDALIB Then
        sMsgOpe = "Retiro Serv Sedalib"
    ElseIf nOperacion = gAhoRetComServEDELNOR Then
        sMsgOpe = "Retiro Serv Edelnor"
    ElseIf nOperacion = gAhoRetEmiChq Then
        sMsgOpe = "Retiro Emisión Cheque"
    ElseIf nOperacion = gAhoRetEmiChqCanjeOP Then
        sMsgOpe = "Ret Emi.Chq Canje OP"
    ElseIf nOperacion = gAhoDctoEmiExt Then
        sMsgOpe = "Ret Emision Ext."
    ElseIf nOperacion = gAhoRetOtrosConceptos Then
        sMsgOpe = "Retiro otros Conceptos"
    ElseIf nOperacion >= 200331 And nOperacion <= 200357 Then
        sMsgOpe = Trim(sOPedesc)
        
    End If
    
    
End If
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.ActualizadUltContactoAho sCuenta, dFecSis

If bInmovilizada Then
    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
        clsCap.AgregaMovCap nMovNro, gAhoEstInmovAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInmovilizadaAct sCuenta
        clsCap.ActualizaInactivaAct sCuenta
    End If
End If

If bInactiva And Not bInmovilizada Then
    clsCap.AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    clsCap.ActualizaInactivaAct sCuenta
End If

clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
If nOperacion = gAhoRetTransf Or nOperacion = gAhoRetEmiChq Then
    clsCap.AgregaMovTranferBanco nMovNro, sCodIF, psCtaBanco, pnMoneda

    If pnMoneda <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis
        
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
    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
End If

'************************
If Not bTransaccion Then clsCap.dbCmact.CommitTrans
bTrans = False

CapCargoCuentaAho = nSaldoCnt

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


'If pbITFAplica And pnITFValor > 0 Then
'    If pbITFAsumido Then
'        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, bNumExtracto
'        clsCap.AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
'        clsCap.AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFAsumido, pnITFValor
'    Else
'        If psITFOperacion = gITFCobroEfectivo Then
'            clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, bNumExtracto
'            clsCap.AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
'            clsCap.AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFCliente, pnITFValor
'        Else
'            clsCap.ActualizaCargoCaptacion sCuenta, nMonto + pnITFValor, nMonto + pnITFValor, nIntGanado, dUltMov, sMovnro, bNumExtracto
'            clsCap.AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
'            clsCap.AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFCliente, pnITFValor
'        End If
'    End If
'Else
'    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, bNumExtracto
'End If


If pbITFAplica And pnITFValor > 0 And pbITFAsumido = False Then
    
        If (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
            nSaldoDisp = nSaldoDisp - pnITFValor
            nSaldoCnt = nSaldoCnt - pnITFValor
            
        End If
End If




If sMsgOpe = "" Then sMsgOpe = clsCap.GetCaptaOperacion(Trim(Str(nOperacion)))

Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocOrdenPago
                If sCodCMAC <> "" Then
                    sTipApe = "RETIRO AHORROS CMACICA"
                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion)
                Else
                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion)
                    
                End If
        End Select
    Else
        If sCodCMAC <> "" Then
            sTipApe = "RETIRO AHORROS CMACICA"
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion)
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion)
        End If
    End If
    
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            End If
        Else
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont, CStr(nOperacion)
            End If
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Description, vbCritical, "Error"
    MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    CapCargoCuentaAho = 0
End Function


Public Function CapCargoCuentaCTS(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF1 As String = "", Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional ByVal SBENPERSLAVDINERO As String = "", _
            Optional sCodIF As String = "", Optional psCtaBanco As String = "", Optional pnMoneda As Moneda = gMonedaNacional, _
            Optional nIntaRet As Double, Optional nPorInta As Double, Optional nPorDisp As Double, _
            Optional nPorIntaCta As Double, Optional nPorDispCta As Double, Optional nMovNroTemp As Long = 0) As Double
            
             

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
Dim sMsgOpe As String, sCodUser As String, nSaldEfectivo As Double
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el cálculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = Round(rsCta("nSaldoDisp"), 2)
nSaldoCnt = Round(rsCta("nSaldo"), 2)
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nSaldRetiro = Round(rsCta("nSaldRetiro"), 2)
nSaldEfectivo = Round(rsCta("nSaldoDisp"), 2)
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


If Not ValidaSaldoCuenta(sCuenta, nMonto, nOperacion) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    CapCargoCuentaCTS = -1
    Exit Function
End If

sGlosa = sGlosa & "Cargo Cuenta = " & sCuenta

'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)

nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
'nIntGanadoRet = GetInteres(nSaldRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanadoRet = GetInteres(nSaldRetiro, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)




dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto
nSaldoDisp = nSaldoDisp - nMonto

clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If nOperacion <> "220303" Then
    clsCap.ActualizaCargoCTS sCuenta, nMonto, nIntGanadoRet
Else
    clsCap.ActualizaCargoCTS sCuenta, nMonto - nIntaRet, nIntGanadoRet
End If

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
    If nOperacion = gCTSRetEfec Then
        sMsgOpe = "Retiro Efectivo"
    ElseIf nOperacion = gCTSRetTransf Then
        sMsgOpe = "Retiro Transferencia Bcos"
    ElseIf nOperacion = "220303" Then
        sMsgOpe = "Retiro LEY 28461"
    End If
End If

clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

If nOperacion = gCTSRetTransf Then
    clsCap.AgregaMovTranferBanco nMovNro, sCodIF, psCtaBanco, pnMoneda

    If pnMoneda <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis
        
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

If nOperacion = "220303" Then
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, 30, nIntaRet
    clsCap.AgregaMovParamCTS nMovNro, 0, nPorInta, nPorDisp
End If

If sCodCMAC <> "" Then
    clsCap.AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If

clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
'ALPA 20081010
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
'End If

'*******************************************************
'ALPA 20081010
nMovNroTemp = nMovNro
clsCap.dbCmact.CommitTrans
bTrans = False
'********************************************************
CapCargoCuentaCTS = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim clsNeg As NCapMantenimiento
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sTipApe = "RETIRO CTS"
Set clsNeg = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsNeg.GetNombreTitulares(sCuenta))

Dim psDispImp As String * 39, psIntaImp As String * 39
Dim psDispCtaImp As String * 39, psIntaCtaImp As String * 39

'Optional nPorIntaCta As Double, Optional nPorDispCta As Double
psDispCtaImp = "Disponible Cuenta: " & Format(nPorDispCta, "##0.0000") & "% " & CStr(nSaldRetiro)
psIntaCtaImp = "Intangible Cuenta: " & Format(nPorIntaCta, "##0.0000") & "% " & CStr(nSaldEfectivo - nSaldRetiro)
'nIntaRet As Double, Optional nPorInta As Double, Optional nPorDisp
psDispImp = "Disponible Retiro: " & Format(nPorDisp, "##0.0000") & "% " & CStr(nMonto - nIntaRet)
psIntaImp = "Intangible Retiro: " & Format(nPorInta, "##0.0000") & "% " & CStr(nIntaRet)


Set clsNeg = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , , , , , , , , , psDispImp, psIntaImp, psDispCtaImp, psIntaCtaImp
    Else
        If sCodCMAC <> "" Then
            sTipApe = "RETIRO CMAC CTS"
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , , , , , , , , , psDispImp, psIntaImp, psDispCtaImp, psIntaCtaImp
        Else
            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , , , , , , , , , psDispImp, psIntaImp, psDispCtaImp, psIntaCtaImp
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo



Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapCargoCuentaCTS = 0
End Function
Private Function GetTasaMinima(ByVal pnProducto As Integer, ByVal pnMoneda As Integer, ByVal pcOrdPag As String, ByVal pcCodAge As String, ByVal pntasatpo, Optional ByVal bOrdPag As Boolean)
Dim ssql As String, RSTEMP As Recordset
Dim conex As DConecta
ssql = "select ntasavalor from captactasas "
ssql = ssql & " where nplazoini=0 and NTASAPROD=" & pnProducto & " AND NTASAMON=" & pnMoneda & "  AND CORDPAG='" & pcOrdPag & "' AND CCODAGE='" & pcCodAge & "' "
ssql = ssql & " AND NTASATPO=" & pntasatpo & " and cordpag='" & IIf(bOrdPag = True, "1", "0") & "'"
Set RSTEMP = New Recordset
Set conex = New DConecta
conex.AbreConexion
RSTEMP.CursorLocation = adUseClient
Set RSTEMP = conex.CargaRecordSet(ssql)
Set RSTEMP.ActiveConnection = Nothing

If RSTEMP.EOF And RSTEMP.BOF Then
    GetTasaMinima = 0
Else
    GetTasaMinima = RSTEMP("ntasavalor")
End If
conex.CierraConexion
RSTEMP.Close
Set RSTEMP = Nothing
Set conex = Nothing
End Function



Public Function GetSaldoCancelacion(ByVal sCuenta As String, ByVal dFecSis As Date, ByVal sCodage As String, _
                Optional ByRef nTasaCancelacion As Double = 0) As Double
Dim nProd As Producto
Dim rsProd As Recordset
Dim clsMant As DCapMantenimiento
Dim nSaldoDisp As Double, nSaldoCnt As Double
Dim nIntAcum As Double, nIntGanado As Double
Dim nTasa As Double
Dim dUltMov As Date
Dim nDiasTranscurridos As Long

nProd = CLng(Mid(sCuenta, 6, 3))
Set clsMant = New DCapMantenimiento
Select Case nProd
    Case gCapAhorros, gCapCTS
        If nProd = gCapAhorros Then
            Set rsProd = clsMant.GetDatosCuentaAho(sCuenta)
        Else
            Set rsProd = clsMant.GetDatosCuentaCTS(sCuenta)
        End If
        Set clsMant = Nothing
        
        nSaldoDisp = rsProd("nSaldoDisp")
        nSaldoCnt = rsProd("nSaldo")
        dUltMov = rsProd("dUltCierre")
        nTasa = rsProd("nTasaInteres")
        nIntAcum = rsProd("nIntAcum")
        nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
        nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
        GetSaldoCancelacion = nSaldoDisp + nIntAcum + nIntGanado
        rsProd.Close
        
       Set rsProd = Nothing
    Case gCapPlazoFijo
        Dim nTasaDias As Double, nIntPagado As Double
        Dim nDias As Long, nPlazo As Long, nPlazoCumplido As Long
        Dim dVencimiento As Date, dApertura As Date, dAperturaReal As Date
        Dim nInt01 As Double, nInt02 As Double, nInt03 As Double
        Dim nTipoTasa As CaptacTipoTasa
        Dim clsCap As DCapMovimientos
        Dim clsDef As nCapDefinicion
        
        
        Set rsProd = clsMant.GetDatosCuentaPF(sCuenta)
        Set clsMant = Nothing
        nSaldoDisp = rsProd("nSaldoDisp")
        nSaldoCnt = rsProd("nSaldo")
      
        
        dApertura = Format(rsProd("dRenovacion"), "dd/MM/yyyy")
        dUltMov = rsProd("dUltCierre")
        nTasa = rsProd("nTasaInteres")
        nIntAcum = rsProd("nIntAcum")
        nIntPagado = rsProd("nIntPag")
        nPlazo = rsProd("nPlazo")
        nTipoTasa = rsProd("nPrdTasaInteres")
        dAperturaReal = Format(rsProd("dApertura"), "dd/MM/yyyy")
        
        dVencimiento = Format(IIf(dApertura = dAperturaReal, DateAdd("d", nPlazo + 1, dApertura), DateAdd("d", nPlazo, dApertura)), "dd/MM/yyyy")
        nPlazoCumplido = DateDiff("d", dApertura, dFecSis)
        Set clsDef = New nCapDefinicion
        'Si la tasa es preferencial, la tasa de cancelacion es la misma tasa
        If nTipoTasa = gCapTasaPreferencial Or nTipoTasa = gCapTasaEspecial Then
            nTasaDias = nTasa
        Else
            nTasaDias = clsDef.GetCapTasaInteres(gCapPlazoFijo, CLng(Mid(sCuenta, 9, 1)), gCapTasaNormal, nPlazoCumplido, nSaldoCnt, sCodage)
        End If
        Set clsDef = Nothing
        nInt01 = 0: nInt02 = 0: nInt03 = 0
        
        If dVencimiento = dFecSis Then
            nInt01 = nIntAcum
            nDias = DateDiff("d", dUltMov, dVencimiento) - 1
            nInt02 = GetInteres(nSaldoDisp + nIntAcum, nTasa, nDias, TpoCalcIntCompuesto)
            nTasaCancelacion = nTasa
        ElseIf dVencimiento > dFecSis Then
            Dim sValor As String
            Set clsCap = New DCapMovimientos
            sValor = clsCap.GetFechaValorChqApePF(sCuenta)
            Set clsCap = Nothing
            If sValor <> "" Then
                dApertura = CDate(sValor)
            End If
            
            nPlazoCumplido = DateDiff("d", dApertura, dFecSis) - 1
            
            If Format(dApertura, "dd/MM/yyyy") = dFecSis Then nPlazoCumplido = 0
            
            If dApertura = dAperturaReal Then
                'nInt01 = nIntAcum
                 'nDias = DateDiff("d", dUltMov, dFecSis) - 1
                nInt01 = 0
                If nPlazoCumplido <= 30 Then
                    nTasaDias = GetTasaMinima(Producto.gCapAhorros, CLng(Mid(sCuenta, 9, 1)), 1, Mid(sCuenta, 4, 2), gCapTasaNormal, True)
                ElseIf nPlazoCumplido > 30 Then
                    nTasaDias = GetTasaMinima(Producto.gCapAhorros, CLng(Mid(sCuenta, 9, 1)), 0, Mid(sCuenta, 4, 2), gCapTasaNormal, False)
                End If
                nInt02 = GetInteres(nSaldoDisp + nInt01, nTasaDias, nPlazoCumplido, TpoCalcIntCompuesto) - nIntPagado
            Else
                nTasaDias = nTasa
               
'*****original sicmact
'               nInt01 = 0
'                nInt02 = GetInteres(nSaldoDisp + nInt01, nTasaDias, nPlazoCumplido, TpoCalcIntCompuesto) - nIntPagado
                nDias = DateDiff("d", dUltMov, dFecSis) - 1
                nInt01 = nIntAcum
                nInt02 = GetInteres(nSaldoDisp + nInt01, nTasaDias, nDias, TpoCalcIntCompuesto)
                
            End If
            nTasaCancelacion = nTasaDias
        Else
            nInt01 = nIntAcum
            
            If DateDiff("d", CDate(dUltMov), CDate(dVencimiento)) >= 1 Then
                nInt02 = GetInteres(nSaldoDisp + nInt01, nTasa, DateDiff("d", CDate(dUltMov), CDate(dVencimiento)) - 1, TpoCalcIntCompuesto)
            Else
                nInt02 = 0
            End If
            nTasaCancelacion = nTasa
            nInt03 = GetInteres(nSaldoDisp + nInt01 + nInt02, nTasa, DateDiff("d", dVencimiento, dFecSis) - 1, TpoCalcIntCompuesto)
        End If
        nIntGanado = nInt01 + nInt02 + nInt03
        GetSaldoCancelacion = nSaldoDisp + nIntGanado
        rsProd.Close
        Set rsProd = Nothing
End Select
End Function

Public Function TieneChequesValorizacion(ByVal sCuenta As String) As Boolean
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
TieneChequesValorizacion = clsCap.TieneChequesValorizacion(sCuenta)
Set clsCap = Nothing
End Function



Public Sub CapCancelaCuentaAho(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal sGlosa As String, ByVal nOperacion As CaptacOperacion, _
        Optional nEstado As CaptacEstado = gCapEstCancelada, Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional ByVal psCodCMAC As String, _
        Optional PbValidaSaldo As Boolean = True, _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional ByVal SBENPERSLAVDINERO As String = "", Optional nMovNroTemp As Long = 0)

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nMonto As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean, bInmovilizada As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

'sMovNro

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")

'Format$(clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, gsCodAge), "#,##0.00")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")

nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
bInmovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion
Set clsCap = New DCapMovimientos
'On Error GoTo ErrGraba

clsCap.dbCmact.BeginTrans
bTrans = True

'***************
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

'If PbValidaSaldo Then
    'If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    '    clsCap.dbCmact.RollbackTrans
    '    Set clsCap = Nothing
    '    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    '    Exit Sub
    'ElseIf TieneChequesValorizacion(sCuenta) Then
    If TieneChequesValorizacion(sCuenta) Then
        clsCap.dbCmact.RollbackTrans
        Set clsCap = Nothing
        MsgBox "La Cuenta Tiene cheques en valorizacion", vbInformation, "Aviso"
        Exit Sub
    End If
 'End If
'**************

nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'nMonto = Round(nSaldoCnt + nIntAcum + nIntGanado, 2)
nMonto = Format(nSaldoCnt + nIntAcum + nIntGanado, "#,##0.00")
dUltMov = DateAdd("d", -1, dFecSis)
If nOperacion = gAhoCancTransfAct Then
    If bInactiva Then nOperacion = gAhoCancTransfInact
    sMsgOpe = "Cancelación Transferencia"
ElseIf nOperacion = gAhoCancAct Then
    If bInactiva Then nOperacion = 200409
    sMsgOpe = "Cancelación Efectivo"
End If
If sGlosa = "" Then sGlosa = "Cancelación Cuenta = " & sCuenta
clsCap.ActualizadUltContactoAho sCuenta, dFecSis
clsCap.ActualizaCargoCaptacion sCuenta, 0, 0, nIntGanado, dUltMov, sMovNro, True
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntAcum + nIntGanado, 0, 0
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, Format(nSaldoCnt + nIntAcum + nIntGanado, "#,##0.00"), 0, 0
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntAcum + nIntGanado

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCuenta, pnITFValor, 0, 0
        clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCuenta, gConcITFAsumido, pnITFValor
    Else
        clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCuenta, pnITFValor, 0, 0
        clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCuenta, gConcITFCliente, pnITFValor
    End If
End If
''ALPA 20081010
''If sPersLavDinero <> "" Then
''    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
''End If
''ALPA 20081010
nMovNroTemp = nMovNro
clsCap.dbCmact.CommitTrans

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String

sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "CANCELACION AHORROS"
Dim clsNMant As NCapMantenimiento
Set clsNMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsNMant.GetNombreTitulares(sCuenta))
Set clsNMant = Nothing

Do
    ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto - (nIntAcum + nIntGanado)), sNomTit, sCuenta, "", 0, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt, psCodCMAC
    
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
        Else
            fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

Exit Sub

ErrGraba:
    'If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Sub CapCancelaCuentaCTS(ByVal sCuenta As String, ByVal sMovNro As String, _
            ByVal sGlosa As String, ByVal nOperacion As CaptacOperacion, Optional sNroDoc As String = "", _
            Optional nDocumento As TpoDoc = TpoDocNotaCargo, Optional sNomAge As String = "", _
            Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional ByVal psCodCMAC As String, Optional pnMovNroTemp As Long = 0)

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nMonto As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nEstado As CaptacEstado
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
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
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True

'***************
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    Exit Sub
ElseIf TieneChequesValorizacion(sCuenta) Then
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "La Cuenta Tiene cheques en valorizacion", vbInformation, "Aviso"
    Exit Sub
End If
'**************

nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
nMonto = Format(nSaldoCnt + nIntAcum + nIntGanado, "#,##0.00")
dUltMov = DateAdd("d", -1, dFecSis)
If nOperacion = gCTSCancEfec Then
    sMsgOpe = "Cancelación Efectivo"
ElseIf nOperacion = gCTSCancTransf Then
    sMsgOpe = "Cancelación Transferencia"
End If
If sGlosa = "" Then
    sGlosa = "Cancelación Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & ".Cancelación Cuenta = " & sCuenta
End If
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, Format(nSaldoDisp + nIntAcum + nIntGanado, "#,##0.00"), 0, 0
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntAcum + nIntGanado
'ALPA 20081010
''If sPersLavDinero <> "" Then
''    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
''End If
pnMovNroTemp = nMovNro
clsCap.dbCmact.CommitTrans

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim clsImp As NCapMantenimiento
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "CANCELACION CTS"
Set clsImp = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsImp.GetNombreTitulares(sCuenta))
Set clsImp = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, nSaldoCnt, sNomTit, sCuenta, sNroDoc, 0, Trim(0), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt, psCodCMAC
    Else
        ImprimeBoletaInteres sTipApe, ImpreCarEsp(sModDep), sCodOpe, nSaldoCnt, sNomTit, sCuenta, "", 0, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntAcum + nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt, psCodCMAC
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub
Public Function CapTransferenciaAhoaNTERIOR(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal sMovNro As String, ByVal rsCtaAbo As Recordset, ByVal sGlosa As String, _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional pnTCC As Double = 0, Optional pnTCV As Double = 0) As Boolean

'Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
'Dim rsCta As Recordset
'Dim nEstado As CaptacEstado
'Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
'Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double, nMontoAbo As Double
'Dim dUltMov As Date
'Dim nSaldoInac As Double
'Dim bInactiva As Boolean, bTrans As Boolean, bActivaCta As Boolean, bInmovilizada As Boolean
'Dim nNumExtracto As Long, nDiasTranscurridos As Long, I As Long
'Dim nExtracto As Long, nMovNro As Long
'Dim sMsgOpe As String, sCtaAbo As String, sCodUser As String
'Dim nMoneda As Moneda
'Dim nOperacion As CaptacOperacion
'Dim dFecSis As Date
'Dim oCon As NConstSistemas
'Set oCon = New NConstSistemas
'
'Dim lnTCC As Double
'Dim lnTCV As Double
'
'Dim clsGen As nTipoCambio
'Set clsGen = New nTipoCambio
'
'dFecSis = CDate(Mid(sMovnro, 7, 2) & "/" & Mid(sMovnro, 5, 2) & "/" & Left(sMovnro, 4))
'
'If pnTCC = 0 Then
'    lnTCC = clsGen.EmiteTipoCambio(dFecSis, TCCompra)
'Else
'    lnTCC = pnTCC
'End If
'
'If pnTCV = 0 Then
'    lnTCV = clsGen.EmiteTipoCambio(dFecSis, TCVenta)
'Else
'    lnTCV = pnTCV
'End If
'
'sCodUser = Right(sMovnro, 4)
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
'clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, True
'sMsgOpe = "Cargo Transferencia"
'clsCap.AgregaMov sMovnro, nOperacion, sGlosa
'nMovNro = clsCap.GetnMovNro(sMovnro)
'clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
'clsCap.ActualizadUltContactoAho sCuenta, dFecSis
'
'If bInmovilizada Then
'    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
'        clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
'        clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovnro
'        clsCap.ActualizaInmovilizadaAct sCuenta
'    End If
'End If
'
'If bInactiva And Not bInmovilizada Then
'    clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
'    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovnro
'    clsCap.ActualizaInactivaAct sCuenta
'End If
'
'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
'clsCap.UltimaActualizacionCuenta sCuenta, sMovnro
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
'Set clsTit = Nothing
'Set oImp = New NCapImpBoleta
'sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
''Inicio de los abonos de cuenta
'Do While Not rsCtaAbo.EOF
'    sCtaAbo = rsCtaAbo("Cuenta")
'    If CLng(Mid(sCtaAbo, 9, 1)) = gMonedaNacional Then
'        nMontoAbo = rsCtaAbo("Monto S/.")
'    Else
'        nMontoAbo = rsCtaAbo("Monto $")
'    End If
'    nOperacion = gAhoTransAbono
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
'    clsCap.ActualizaAbonoCaptacion sCtaAbo, nMontoAbo, nMontoAbo, nIntGanado, dUltMov, sMovnro, True
'    sMsgOpe = "Abono Transferencia"
'    If sGlosa = "" Then sGlosa = "Abono Transferencia Cuenta = " & sCtaAbo
'    clsCap.AgregaMovCap nMovNro, nOperacion, sCtaAbo, nMontoAbo, nSaldoDisp, nSaldoCnt
'    clsCap.ActualizadUltContactoAho sCtaAbo, dFecSis
'
'    If bInmovilizada Then
'        If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
'            clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCtaAbo, gConcCapital, nSaldoInac
'            clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovnro
'            clsCap.ActualizaInmovilizadaAct sCtaAbo
'        End If
'    End If
'
'    If bInactiva And Not bInmovilizada Then
'        clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCtaAbo, gConcCapital, nSaldoInac
'        clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovnro
'        clsCap.ActualizaInactivaAct sCtaAbo
'    End If
'
'    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCtaAbo, gConcCapital, nMontoAbo
'    clsCap.UltimaActualizacionCuenta sCtaAbo, sMovnro
'    'Impresión de la boleta
'    sTipDep = IIf(Mid(sCtaAbo, 9, 1) = "1", "SOLES", "DOLARES")
'    sCodOpe = Trim(nOperacion)
'    sModDep = sMsgOpe
'    sTipApe = "TRANSFERENCIA AHORROS"
'    Set clsTit = New NCapMantenimiento
'    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCtaAbo))
'    Set clsTit = Nothing
'    sBoletas = sBoletas & Chr$(12)
'    If Mid(sCuenta, 9, 1) = Mid(sCtaAbo, 9, 1) Then
'        sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
'    Else
'        If Mid(sCtaAbo, 9, 1) = "1" Then
'            sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , "Imp U$. " & Format(rsCtaAbo.Fields(3), "#,##0.00") & " A TCC:" & Format(lnTCC, "#,##0.000"), , dFecSis, sNomAge, sCodUser, sLpt, True, False)
'        Else
'            sBoletas = sBoletas & oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", 1, nSaldoCnt, True, , , , , , , "Imp S/. " & Format(rsCtaAbo.Fields(2), "#,##0.00") & " A TCV:" & Format(lnTCV, "#,##0.000"), , dFecSis, sNomAge, sCodUser, sLpt, True, False)
'        End If
'    End If
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
End Function
'ALPA 20081009****************************************************************************
'''Public Function CapTransferenciaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
'''            ByVal sMovNro As String, ByVal rsCtaAbo As Recordset, ByVal sGlosa As String, _
'''            Optional sNomAge As String = "", Optional sLpt As String = "", _
'''            Optional sPersLavDinero As String = "", Optional pnTCC As Double = 0, Optional pnTCV As Double = 0, _
'''            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional ByVal SBENPERSLAVDINERO As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "", _ ') As Boolean
        Public Function CapTransferenciaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal sMovNro As String, ByVal rsCtaAbo As Recordset, ByVal sGlosa As String, _
        Optional sNomAge As String = "", Optional sLpt As String = "", _
        Optional sPersLavDinero As String = "", Optional pnTCC As Double = 0, Optional pnTCV As Double = 0, _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional ByVal SBENPERSLAVDINERO As String = "", Optional pnMonNro As Long = 0) As Boolean

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double, nMontoAbo As Double
Dim dUltMov As Date
Dim nSaldoInac As Double
Dim bInactiva As Boolean, bTrans As Boolean, bActivaCta As Boolean, bInmovilizada As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCtaAbo As String, sCodUser As String
Dim nMoneda As Moneda
Dim nOperacion As CaptacOperacion
Dim dFecSis As Date
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Dim i As Integer


Dim lnTCC As Double
Dim lnTCV As Double

Dim clsGen As nTipoCambio
Set clsGen = New nTipoCambio

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))

If pnTCC = 0 Then
    lnTCC = clsGen.EmiteTipoCambio(dFecSis, TCCompra)
Else
    lnTCC = pnTCC
End If

If pnTCV = 0 Then
    lnTCV = clsGen.EmiteTipoCambio(dFecSis, TCVenta)
Else
    lnTCV = pnTCV
End If

sCodUser = Right(sMovNro, 4)
bTrans = False
bActivaCta = False
nMoneda = CLng(Mid(sCuenta, 9, 1))
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
clsCap.dbCmact.BeginTrans
bTrans = True
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox "Cuenta NO Posee SALDO SUFICIENTE", vbInformation, "Aviso"
    CapTransferenciaAho = False
    Exit Function
End If
sGlosa = sGlosa & "Cargo Transferencia Cuenta = " & sCuenta
nOperacion = gAhoTransCargo
If bInactiva Then
    nSaldoInac = nSaldoCnt
End If
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto
clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto
clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

sMsgOpe = "Cargo Transferencia"
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
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

'**********PARA AGREGAR ITF***********
If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        'clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, True
        clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCuenta, pnITFValor, 0, 0
        clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCuenta, gConcITFAsumido, pnITFValor
        
    Else
        If psItfOperacion = gITFCobroEfectivo Then
           ' clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovnro, True
            clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCuenta, gConcITFCliente, pnITFValor
        Else
            clsCap.ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, True
            clsCap.AgregaMovCap nMovNro, gITFCobroCargo, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
            clsCap.AgregaMovCapDet nMovNro, gITFCobroCargo, sCuenta, gConcITFCliente, pnITFValor
            nSaldoDisp = nSaldoDisp - pnITFValor
            nSaldoCnt = nSaldoCnt - pnITFValor
        End If
    
    End If
End If
'**********FIN ITF*************

clsCap.UltimaActualizacionCuenta sCuenta, sMovNro


Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim sBoletas As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO AHORROS"
Dim clsTit As NCapMantenimiento
Dim oImp As NCapImpBoleta
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
Set clsTit = Nothing
Set oImp = New NCapImpBoleta
Dim oPrevio As previo.clsprevio

  i = 0
  Do
        i = i + 1

        Set oPrevio = New previo.clsprevio
        sBoletas = ""
        sBoletas = oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado + nIntAcum, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
        
            oPrevio.PrintSpool sLpt, sBoletas, False
        


        '*******BOLETAS ITF CTA CARGA
        If pbITFAplica And pnITFValor > 0 Then
                If pbITFAsumido Then
        '            If psItfOperacion = gITFCobroEfectivo Then
                        fgITFImprimeBoletaTrans sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
        '            Else
        '                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovnro, 7, 2) & "/" & Mid(sMovnro, 5, 2) & "/" & Left(sMovnro, 4)), sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt
        '            End If
                Else
                    If psItfOperacion = gITFCobroEfectivo Then
                        fgITFImprimeBoletaTrans sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
                    Else
                        fgITFImprimeBoletaTrans sNomTit, pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
                    End If
                End If
        End If
        
        '*******BOLETAS ITF
        

   Loop Until MsgBox("Reimprimir Boletas de CTa a Cargar en Transferencia", vbYesNo + vbQuestion, "AVISO") = vbNo


'Inicio de los abonos de cuenta
Dim CANTABO As Integer
CANTABO = 0

Do While Not rsCtaAbo.EOF

    Dim nMonITF As Double, sAboOpeITF As String

    sCtaAbo = rsCtaAbo("Cuenta")
    
    If CLng(Mid(sCtaAbo, 9, 1)) = gMonedaNacional Then
        If rsCtaAbo("fItf") = "S" And rsCtaAbo("TipoItf") = "C" Then
            nMontoAbo = rsCtaAbo("Monto S/.")
            nMonITF = rsCtaAbo("ITF S/.")
            sAboOpeITF = gITFCobroCargo
            
        ElseIf (rsCtaAbo("fItf") = "S" And rsCtaAbo("TipoItf") = "E") Or rsCtaAbo("TipoItf") = "A" Then
            nMontoAbo = rsCtaAbo("Monto S/.")
            nMonITF = rsCtaAbo("ITF S/.")
            sAboOpeITF = gITFCobroEfectivo
            
        ElseIf rsCtaAbo("fItf") = "N" Then
            nMontoAbo = rsCtaAbo("Monto S/.")
            nMonITF = 0
            sAboOpeITF = gITFCobroEfectivo
            
        End If
    Else
        If rsCtaAbo("fItf") = "S" And rsCtaAbo("TipoItf") = "C" Then
            nMontoAbo = rsCtaAbo("Monto $")
            nMonITF = rsCtaAbo("ITF $")
            sAboOpeITF = gITFCobroCargo
        ElseIf (rsCtaAbo("fItf") = "S" And rsCtaAbo("TipoItf") = "E") Or rsCtaAbo("TipoItf") = "A" Then
            nMontoAbo = rsCtaAbo("Monto $")
            nMonITF = rsCtaAbo("ITF $")
            sAboOpeITF = gITFCobroEfectivo
        ElseIf rsCtaAbo("fItf") = "N" Then
            nMontoAbo = rsCtaAbo("Monto $")
            nMonITF = 0
            sAboOpeITF = gITFCobroEfectivo
        End If
        
        
    End If
    
    nOperacion = gAhoTransAbono
    Set clsMant = New DCapMantenimiento
    Set rsCta = clsMant.GetDatosCuentaAho(sCtaAbo)
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
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMontoAbo
    clsCap.ActualizaSaldoAnteriorAho sCtaAbo, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMontoAbo
    clsCap.ActualizaAbonoCaptacion sCtaAbo, nMontoAbo, nMontoAbo, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Abono Transferencia"
    If sGlosa = "" Then sGlosa = "Abono Transferencia Cuenta = " & sCtaAbo
    clsCap.AgregaMovCap nMovNro, nOperacion, sCtaAbo, nMontoAbo, nSaldoDisp, nSaldoCnt
    clsCap.ActualizadUltContactoAho sCtaAbo, dFecSis

    If bInmovilizada Then
        If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
            clsCap.AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCtaAbo, gConcCapital, nSaldoInac
            clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovNro
            clsCap.ActualizaInmovilizadaAct sCtaAbo
        End If
    End If
    
    If bInactiva And Not bInmovilizada Then
        clsCap.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCtaAbo, gConcCapital, nSaldoInac
        clsCap.ActualizaEstadoCuenta sCtaAbo, gCapEstActiva, dFecSis, sMovNro
        clsCap.ActualizaInactivaAct sCtaAbo
    End If
    
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCtaAbo, gConcCapital, nMontoAbo
    
    '**********PARA AGREGAR ITF***********
    If pbITFAplica And nMonITF > 0 Then
            If rsCtaAbo("TipoItf") = "A" Then
                'clsCap.ActualizaAbonoCaptacion sCtaAbo, nMontoAbo, nMontoAbo, nIntGanado, dUltMov, sMovnro, True
                clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCtaAbo, nMonITF, 0, 0
                clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCtaAbo, gConcITFAsumido, nMonITF
            Else
                If rsCtaAbo("TipoItf") = "E" Then
                    'clsCap.ActualizaAbonoCaptacion sCtaAbo, nMontoAbo, nMontoAbo, nIntGanado, dUltMov, sMovnro, True
                    clsCap.AgregaMovCap nMovNro, gITFCobroEfectivo, sCtaAbo, nMonITF, nSaldoDisp, nSaldoCnt
                    clsCap.AgregaMovCapDet nMovNro, gITFCobroEfectivo, sCtaAbo, gConcITFCliente, nMonITF
                Else
                    clsCap.ActualizaAbonoCaptacion sCtaAbo, -nMonITF, -nMonITF, nIntGanado, dUltMov, sMovNro, True
                    clsCap.AgregaMovCap nMovNro, gITFCobroCargo, sCtaAbo, nMonITF, nSaldoDisp - nMonITF, nSaldoCnt - nMonITF
                    clsCap.AgregaMovCapDet nMovNro, gITFCobroCargo, sCtaAbo, gConcITFCliente, nMonITF
                    
                    nSaldoDisp = nSaldoDisp - nMonITF
                    nSaldoCnt = nSaldoCnt - nMonITF
                    
                End If
            End If

    End If


    '**********FIN ITF*************
    
    
    clsCap.UltimaActualizacionCuenta sCtaAbo, sMovNro
    
    
    CANTABO = CANTABO + 1
    
    
    'Impresión de la boleta
    sTipDep = IIf(Mid(sCtaAbo, 9, 1) = "1", "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "TRANSFERENCIA AHORROS"
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCtaAbo))
    Set clsTit = Nothing
    'sBoletas = sBoletas & Chr$(12)
    'sBoletas = "" & Chr$(12)
    
    i = 0
    Do
        i = i + 1
    
        If Mid(sCuenta, 9, 1) = Mid(sCtaAbo, 9, 1) Then
            sBoletas = oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, True, False)
        Else
            If Mid(sCtaAbo, 9, 1) = "1" Then
                sBoletas = oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", nExtracto, nSaldoCnt, True, , , , , , , "Imp U$. " & Format(rsCtaAbo.Fields(3), "#,##0.00") & " A TCC:" & Format(lnTCC, "#,##0.000"), , dFecSis, sNomAge, sCodUser, sLpt, True, False)
            Else
                sBoletas = oImp.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAbo), sNomTit, sCtaAbo, "", nSaldoDisp, nIntGanado + nIntAcum, "", nExtracto, nSaldoCnt, True, , , , , , , "Imp S/. " & Format(rsCtaAbo.Fields(2), "#,##0.00") & " A TCV:" & Format(lnTCV, "#,##0.000"), , dFecSis, sNomAge, sCodUser, sLpt, True, False)
            End If
        End If
        Set oPrevio = New previo.clsprevio
        oPrevio.PrintSpool sLpt, sBoletas, False
                            
          
        '*******BOLETAS ITF CTA ABONA
            If pbITFAplica And pnITFValor > 0 Then
                If rsCtaAbo("TipoItf") = "A" Then
        '            If psItfOperacion = gITFCobroEfectivo Then
                        fgITFImprimeBoletaTrans sNomTit, nMonITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCtaAbo, 9, 1), sCtaAbo, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
        '            Else
        '                fgITFImprimeBoleta sNomTit, nMonITF, sGlosa, CDate(Mid(sMovnro, 7, 2) & "/" & Mid(sMovnro, 5, 2) & "/" & Left(sMovnro, 4)), sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt
        '            End If
                Else
                    If rsCtaAbo("TipoItf") = "E" Then
                        fgITFImprimeBoletaTrans sNomTit, nMonITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCtaAbo, 9, 1), sCtaAbo, , False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
                    Else
                        fgITFImprimeBoletaTrans sNomTit, nMonITF, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 2, Mid(sCtaAbo, 9, 1), sCtaAbo, , False, sNomAge, sCodUser, False, nSaldoDisp, nSaldoCnt, i
                    End If
                End If
            End If
        
        '*******BOLETAS ITF
        
        
        
   Loop Until MsgBox("Desea reimprimir boletas de Cta a Abonar " & CStr(CANTABO) & " en transferencia", vbYesNo + vbQuestion, "Aviso") = vbNo
    
    rsCtaAbo.MoveNext
Loop

''If sPersLavDinero <> "" Then
''    'ALPA 20081009*****************************************************************
''    'clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
''    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO, nTipoREU, nMontoAcu, sOrigen
''    '******************************************************************************
''End If
'ALPA 20081010*********************************************
pnMonNro = nMovNro
clsCap.dbCmact.CommitTrans
CapTransferenciaAho = True
Set clsCap = Nothing
'oPrevio.Show sBoletas, ""

Set oImp = Nothing
Set oPrevio = Nothing
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsMant = Nothing
    Set oImp = Nothing
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapTransferenciaAhoCambios = False
End Function

Public Function GetCapInteresAlaFecha(ByVal nProducto As Producto, ByVal nMonto As Double, _
        ByVal nTasa As Double, ByVal dUltMov As Date, ByVal nIntAcum As Double, _
        ByVal dFecSis As Date) As Double
Dim nIntGanado As Double
Dim nDias As Long
nDias = DateDiff("d", dUltMov, dFecSis) - 1
If nDias < 0 Then nDias = 0
Select Case nProducto
    Case gCapAhorros, gCapCTS
        GetCapInteresAlaFecha = nIntAcum + GetInteres(nMonto, nTasa, nDias, TpoCalcIntSimple)
    Case gCapPlazoFijo
        GetCapInteresAlaFecha = nIntAcum + GetInteres(nMonto + nIntAcum, nTasa, nDias, TpoCalcIntCompuesto)
End Select
End Function

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgecod As String, Optional nTipoBus As Integer = 0) As Recordset
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
Set GetMovExtorno = clsCap.GetMovExtorno(sDatoBus, dFecha, nOperacion, psAgecod, nTipoBus)
Set clsCap = Nothing
End Function

Public Function GetMovExtornoGiro(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgecod As String, Optional nTipoBus As Integer = 0) As Recordset
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
Set GetMovExtornoGiro = clsCap.GetMovExtornoGiro(sDatoBus, dFecha, nOperacion, psAgecod, nTipoBus)
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
            Optional sCodIF As String = "", Optional sNomAge As String, Optional sLpt As String = "", _
            Optional ByVal psCodCMAC As String, Optional ByVal pnITFValor As Double = 0, _
            Optional ByVal nITFOperacion As CaptacOperacion, Optional ByVal nITFConcepto As CaptacConcepto) As Double

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
Dim nITFDescuento As Double

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)

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
nITFDescuento = 0
If pnITFValor > 0 And nITFConcepto = gConcITFCliente And (nITFOperacion = gITFCobroCargo Or nITFOperacion = gITFCobroCMACCargo) Then
    nITFDescuento = pnITFValor
End If

If sNroDoc = "" Then
    If Not ValidaSaldoCuenta(sCuenta, nMonto - nITFDescuento) Then
        Set clsCap = Nothing
        MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
        CapExtornoAbonoAho = 0
        Exit Function
    End If
End If
sGlosa = sGlosa & ". Extorno Abono Cuenta = " & sCuenta
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
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto - nITFDescuento, 0, nIntGanado, dUltMov, sMovNro, True
        clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado, , sCuenta
        sMsgOpe = "Ext. Depósito Chq."
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto - nITFDescuento, nMonto - nITFDescuento, nIntGanado, dUltMov, sMovNro, True
    If nOperacion = gAhoExtTransAbono Then
        sMsgOpe = "Ext. Depósito Transferencia"
    ElseIf nOperacion = gAhoExtDepEfec Then
        sMsgOpe = "Ext. Depósito Efectivo"
    ElseIf nOperacion = gAhoExtDepPagServEdelnor Then
        sMsgOpe = "Ext. Dep. Pago Edelnor"
    ElseIf nOperacion = gAhoExtDepOtrosConceptos Then
        sMsgOpe = "Ext. Dep. Otros Conc."
    ElseIf nOperacion = gAhoExtDepDevCredPersonales Then
        sMsgOpe = "Ext. Dep. Dev.Cred.Pers"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

If pnITFValor > 0 And nITFConcepto = gConcITFCliente Then
    clsCap.AgregaMovCap nMovNro, gITFCobroCargoExt, sCuenta, pnITFValor, nSaldoDisp + pnITFValor, nSaldoCnt + pnITFValor
    clsCap.AgregaMovCapDet nMovNro, gITFCobroCargoExt, sCuenta, gConcITFCliente, pnITFValor
End If

clsCap.UltimaActualizacionCuenta sCuenta, sMovNro

clsCap.dbCmact.CommitTrans
bTrans = False
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
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If

    If pnITFValor > 0 Then
        If nITFConcepto = gConcITFAsumido Then
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Abono Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Abono Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            End If
        Else
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Abono Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Abono Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            End If
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    CapExtornoAbonoAho = 0
End Function
Public Sub CapExtornaApeEspecial(ByVal sCtaCod As String, ByVal sMovNro As String)
Dim ssql As String, RSTEMP As Recordset
Dim conex As DConecta
Set RSTEMP = New Recordset
Set conex = New DConecta


conex.AbreConexion
ssql = "SELECT NNUMSOLICITUD,NPRODUCTO,CPERSCOD,NMONEDA,NPLAZO,NTASA,NMONTO FROM CapTasaEspecial WHERE NESTADO=2 AND  CCTACOD='" & sCtaCod & "'"

RSTEMP.CursorLocation = adUseClient
RSTEMP.Open ssql, conex.ConexionActiva, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not (RSTEMP.EOF Or RSTEMP.BOF) Then

    ssql = "INSERT INTO CAPTASAESPECIAL"
    ssql = ssql & "VALUES(" & RSTEMP!nNumSolicitud & ",'" & RSTEMP!cPersCod & "'," & RSTEMP!nProducto & "," & RSTEMP!nMoneda & ",1,'" & sMovNro & "', " & RSTEMP!nTasa & ",'EXTORNO ESTADO 2','" & sCuenta & "'," & RSTEMP!nMonto & "," & RSTEMP!nPlazo & " )"
 
    conex.ConexionActiva.Execute ssql
    RSTEMP.Close
End If
RSTEMP.ActiveConnection = Nothing
conex.CierraConexion
Set RSTEMP = Nothing
Set conex = Nothing
End Sub



Public Function CapExtornoApertura(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional bDocumento As Boolean = False, Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String, _
            Optional ByVal pnITFValor As Double = 0, Optional ByVal nITFOperacion As CaptacOperacion, Optional ByVal nITFConcepto As CaptacConcepto) As String

Dim clsCap As DCapMovimientos


Dim sMsgProd As String, sCodUser As String, sMovNroBus As String, sCodUserBus As String
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

Dim clsMov As DMov
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)

clsCap.dbCmact.BeginTrans
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)

clsCap.ActualizaEstadoCuenta sCuenta, gCapEstAnulada, dFecSis, sMovNro
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, 0, 0
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

'Dim oClsMant As NCapMantenimiento
Dim RSTMP As Recordset, nTipoTasaInt As Integer

'Set oClsMant = New NCapMantenimiento

'Set rstmp = oClsMant.GetDatosCuenta(sCuenta)
'nTipoTasaInt = rstmp!nPrdTasaInteres

Dim oServ1 As DCapServicios, nNumsol As Integer
Set oServ1 = New DCapServicios
 
  Set RSTMP = oServ1.GetEstadoTPCta(sCuenta)
   nNumsol = 0
   If Not RSTMP.EOF Then
     nNumsol = RSTMP!nNumSolicitud
   End If
  
   If nNumsol <> 0 Then
     oServ1.AgregaCapTasaEspecial nNumsol, RSTMP!cPersCod, RSTMP!nProducto, RSTMP!nMoneda, 3, sMovNro, RSTMP!nTasa, "EXTORNO DE APERTURA CON SOLICITUD DE TASA ESPECIAL ", RSTMP!nMonto, sCuenta, RSTMP!nPlazo, 2, RSTMP!bPermanente
   End If
        


Set RSTMP = Nothing
Set oServ1 = Nothing



If pnITFValor > 0 And nITFOperacion = gITFCobroCargo And nITFConcepto = gConcITFCliente Then
    clsCap.AgregaMovCap nMovNro, gITFCobroCargoExt, sCuenta, pnITFValor, 0, 0
    clsCap.AgregaMovCapDet nMovNro, gITFCobroCargoExt, sCuenta, gConcITFCliente, pnITFValor
End If

'Call CapExtornaApeEspecial(sCuenta, sMovNro)

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
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", 0, 0, "", 1, 0, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", 0, 0, "", 1, 0, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If
    
    If pnITFValor > 0 Then
        If nITFConcepto = gConcITFAsumido Then
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Apertura Cta = " & sCuenta, dFecSis, sLpt, 0, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Apertura Cta = " & sCuenta, dFecSis, sLpt, 0, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            End If
        Else
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Apertura Cta = " & sCuenta, dFecSis, sLpt, 0, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Apertura Cta = " & sCuenta, dFecSis, sLpt, 0, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            End If
        End If
    End If

Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

End Function

Public Function CapExtornoAperturaLote(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional bDocumento As Boolean = False, Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String) As String

Dim clsCap As DCapMovimientos
Dim sMsgProd As String, sCodUser As String
Dim nProd As Producto
Dim nMovNro As Long
Dim dFecSis As Date
Dim rsCtas As ADODB.Recordset
Set rsCtas = New ADODB.Recordset

Dim nitf As Currency

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
Set rsCtas = clsCap.GetApeLoteCuentas(nMovNroBus)
rsCtas.MoveFirst


Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)


clsCap.dbCmact.BeginTrans
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)

While Not rsCtas.EOF
    sCuenta = rsCtas!cCtaCod
    nMonto = rsCtas!nMonto
    nitf = rsCtas!ITF
    clsCap.ActualizaEstadoCuenta sCuenta, gCapEstAnulada, dFecSis, sMovNro
    clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, 0, 0
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
    
    If nitf > 0 Then
        clsCap.AgregaMovCap nMovNro, gITFCobroCargoExt, sCuenta, nitf, 0, 0
        clsCap.AgregaMovCapDet nMovNro, gITFCobroCargoExt, sCuenta, gConcITFCliente, nitf
    End If
    
    clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
    rsCtas.MoveNext
Wend

If bDocumento Then
    If nTipoDoc = TpoDocCheque Then
        clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
    End If
End If

clsCap.dbCmact.CommitTrans

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim lsBoletas As String

Do
  rsCtas.MoveFirst
  While Not rsCtas.EOF
    sCuenta = rsCtas!cCtaCod
    nMonto = rsCtas!nMonto
    sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = "Extorno Apertura Lote"
    sTipApe = sMsgProd
    Dim clsTit As NCapMantenimiento
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
    Set clsTit = Nothing
    If bDocumento Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", 0, 0, "", 1, 0, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", 0, 0, "", 1, 0, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If
    rsCtas.MoveNext
  Wend
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

End Function
Public Function GetITFMOV(ByVal nMovNroBus As Long) As Double
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
 GetITFMOV = clsCap.GetITFMOV(nMovNroBus)
Set clsCap = Nothing
End Function
Public Function CapExtornoCargoAho(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional ByVal psCodCMAC As String, Optional ByVal pnITFValor As Double = 0, _
            Optional ByVal nITFOperacion As CaptacOperacion, Optional ByVal nITFConcepto As CaptacConcepto) As Double

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
Dim nITFMonto As Double


Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)


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
If nOperacion <> gAhoExtRetComOrdPagDev Then
    clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
    nITFMonto = GetITFMOV(nMovNroBus)
    If pnITFValor > 0 And nITFConcepto = gConcITFCliente Then
        nITFMonto = pnITFValor
    End If
    clsCap.ActualizaAbonoCaptacion sCuenta, nMonto + nITFMonto, nMonto + nITFMonto, nIntGanado, dUltMov, sMovNro, True
End If

If sNroDoc = "" Then
    If nOperacion = gAhoExtRetEfec Then
        sMsgOpe = "Ext. Retiro Efectivo"
    ElseIf nOperacion = gAhoExtRetConsultaSaldos Then
        sMsgOpe = "Ext. Ret.Cons.Saldos"
    ElseIf nOperacion = gAhoExtRetPorteCargoCuenta Then
        sMsgOpe = "Ext. Ret.Porte N/C"
    ElseIf nOperacion = gAhoExtRetComVentaBases Then
        sMsgOpe = "Ext. Ret. Com.Venta Bases"
    ElseIf nOperacion = gAhoExtTransCargo Then
        sMsgOpe = "Ext. Ret.Transferencia"
    ElseIf nOperacion = gAhoExtDctoEmiExt Then
        sMsgOpe = "Ext. Ret.Emi.Ext."
    ElseIf nOperacion = gAhoExtDctoEmiOP Then
        sMsgOpe = "Ext. Ret.Emi.OP"
        
    ElseIf nOperacion = gAhoExtRetComOrdPagDev Then
        sMsgOpe = "Ext. Retiro Com. OP Dev."
    ElseIf nOperacion = gAhoExtRetComServEdelnor Then
        sMsgOpe = "Ext. Ret.Com.Edelnor"
    ElseIf nOperacion = gAhoExtRetComServEdelnor Then
        sMsgOpe = "Ext. Ret. Otros Conc."
    End If
    
Else
    If nTipoDoc = TpoDocOrdenPago Then
        If nOperacion = gAhoExtRetOPCanje Then
            sMsgOpe = "Ext. Retiro OP Canje"
        ElseIf nOperacion = gAhoExtRetComOrdPagDev Then
            sMsgOpe = "Ext. Retiro Com. OP Dev."
        ElseIf nOperacion = gAhoExtRetFondoFijo Then
            sMsgOpe = "Ext. Retiro Fondo Fijo OP "
        Else
            sMsgOpe = "Ext. Retiro OP"
        End If
       ' clsCap.AgregaOrdenPago sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
       clsCap.EliminaOrdenPago sCuenta, sNroDoc
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If


clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

If pnITFValor > 0 And nITFConcepto = gConcITFCliente And (nITFOperacion = gITFCobroCargo Or nITFOperacion = gITFCobroCMACCargo) Then
    clsCap.AgregaMovCap nMovNro, gITFCobroCargoExt, sCuenta, pnITFValor, nSaldoDisp + pnITFValor, nSaldoCnt + pnITFValor
    clsCap.AgregaMovCapDet nMovNro, gITFCobroCargoExt, sCuenta, gConcITFCliente, pnITFValor
End If

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

If sModDep = "" Then sModDep = clsCap.GetCaptaOperacion(Trim(Str(nOperacion)))

Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If

    If pnITFValor > 0 Then
        If nITFConcepto = gConcITFAsumido Then
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Cargo Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Cargo Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            End If
        Else
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Cargo Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, "Extorno Cargo Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            End If
        End If
    End If

Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCargoAho", Err.Description
    CapExtornoCargoAho = 0
End Function

Public Function CapExtornoCargoAhoOPCertificada(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            ByVal sNroDoc As String, Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String = "") As Double

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

Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)



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
clsCap.AgregaOrdenPago sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
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
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
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
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , gsCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoCargoAhoOPCertificada = 0
End Function

Public Function CapExtornoCargoCTS(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String) As Double

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
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)




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

nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
nIntGanadoRet = GetInteres(nSaldRetiro, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)

'nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
'nIntGanadoRet = GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto

If nOperacion <> "250303" Then
    nSaldoRetiro = nSaldoRetiro + nMonto
Else

  Dim nIntaRet As Double
    nIntaRet = clsCap.GetIntaRetMov(nMovNroBus)
    nSaldoRetiro = nSaldoRetiro + (nMonto - nIntaRet)
    
End If

clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo + nIntGanadoRet

If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
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
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCargoAho", Err.Description
    CapExtornoCargoCTS = 0
End Function

Public Function CapExtornoCancelacion(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String, _
            Optional ByVal nITFValor As Double = 0, Optional ByVal nITFOperacion As CaptacOperacion, _
            Optional ByVal nITFConcepto As CaptacConcepto) As Double

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

Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)



dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nEstado = rsCta("nPrdEstado")
nIntAcum = rsCta("nIntAcum")
nExtracto = rsCta("nTransacc")
rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing

'Inicia la transaccion
On Error GoTo ErrGraba
Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = "Extorno Cancelacion Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & "Extorno Cancelación Cuenta = " & sCuenta
End If
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
If sNroDoc = "" Then
    sMsgOpe = "Ext. Canc. Efectivo"
Else
    If nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Canc. NC"
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
bTrans = False
CapExtornoCancelacion = nSaldoCnt
Set clsCap = Nothing

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe

If CLng(Mid(sCuenta, 6, 3)) = gCapAhorros Then
    sTipApe = "EXTORNO AHORROS"
ElseIf CLng(Mid(sCuenta, 6, 3)) = gCapPlazoFijo Then
    sTipApe = "EXTORNO PLAZO FIJO"
ElseIf CLng(Mid(sCuenta, 6, 3)) = gCapCTS Then
    sTipApe = "EXTORNO CTS"
End If

Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If
    
    If nITFValor > 0 Then
        If nITFConcepto = gConcITFAsumido Then
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, nITFValor, sTipApe & " Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, nITFValor, sTipApe & " Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            End If
        Else
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, nITFValor, sTipApe & " Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, nITFValor, sTipApe & " Cuenta = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            End If
        End If
    End If
    
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Set clsTit = Nothing
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCancelacionAho", Err.Description
    CapExtornoCancelacion = 0
End Function

Public Sub CapExtornoTransferenciaAho(ByVal nMovNroBus As Long, ByVal sMovNro As String, _
            ByVal sGlosa As String, Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String)
Dim clsCap As DCapMovimientos
Dim rsOpe As Recordset, rsItf As Recordset
Dim nOperacion As CaptacOperacion, nOpeExt As CaptacOperacion
Dim nSaldo As Double
Dim sCuenta As String
Dim nMonto As Double
Dim nMovNro As Long
Dim nITFValor As Double, nITFOP As String, nITFConcepto As Integer
On Error GoTo ErrGraba

Dim clsMov As DMov, sCodUserBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(MovNroBus)
sCodUserBus = Right(sMovNroBus, 4)

Set clsCap = New DCapMovimientos
Set rsOpe = clsCap.GetMovTransferencia(nMovNroBus)
Set clsCap = Nothing
'*******validar saldos de cuenta a extornar abonos
If TransValidaSaldos(rsOpe) = False Then Exit Sub
rsOpe.MoveFirst
If Not (rsOpe.EOF And rsOpe.BOF) Then
    Do While Not rsOpe.EOF
        nOperacion = CLng(rsOpe("cOpeCod"))
        nMonto = rsOpe("nMonto")
        sCuenta = rsOpe("cCtaCod")
        nMovNro = rsOpe("nMovNro")
        nITFValor = 0
        nITFOP = ""
        nITFConcepto = 0
        Set clsCap = New DCapMovimientos
        Set rsItf = clsCap.GetMovTransferenciaITF(nMovNro, sCuenta)
            If Not rsItf.EOF Then
                nITFValor = rsItf("nMonto")
                nITFOP = rsItf("copecod")
                nITFConcepto = rsItf("nconceptocod")
            End If
        
        If nOperacion = gAhoTransCargo Then
            If Abs(nITFValor) > 0 Then
              nSaldo = CapExtornoCargoAho(nMovNro, gAhoExtTransCargo, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt, psCodCMAC, nITFValor, nITFOP, nITFConcepto)
            Else
              nSaldo = CapExtornoCargoAho(nMovNro, gAhoExtTransCargo, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt, psCodCMAC)
            End If
        Else
            
            If Abs(nITFValor) > 0 Then
                nSaldo = CapExtornoAbonoAho(nMovNro, gAhoExtTransAbono, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt, psCodCMAC, nITFValor, nITFOP, nITFConcepto)
            Else
                nSaldo = CapExtornoAbonoAho(nMovNro, gAhoExtTransAbono, sCuenta, sMovNro, sGlosa, nMonto, , , , sNomAge, sLpt, psCodCMAC)
            End If
            
        End If
        rsOpe.MoveNext
    Loop
End If
Exit Sub
ErrGraba:
    Err.Raise Err.Number, "CapExtornoTransferenciaAho", Err.Description
End Sub

Public Function TransValidaSaldos(ByVal rs As Recordset) As Boolean
Dim nITFDescuento As Double, pnITFValor As Double, nITFConcepto As Integer, nITFOperacion As String
Dim nMonto As Double, nMovNro As Long, sCuenta As String, sCopecod As String
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos


TransValidaSaldos = True
rs.MoveFirst
While Not rs.EOF

        nMovNro = rs!nMovNro
        sCuenta = rs!cCtaCod
        sCopecod = rs!copecod

        pnITFValor = 0
        nITFConcepto = 0
        nITFOperacion = ""
        
        

        Set rsItf = clsCap.GetMovTransferenciaITF(nMovNro, sCuenta)
        If Not rsItf.EOF Then
                pnITFValor = rsItf("nMonto")
                nITFOperacion = rsItf("copecod")
                nITFConcepto = rsItf("nconceptocod")
        End If
        sCuenta = rs!cCtaCod
        nMonto = rs!nMonto

        nITFDescuento = 0
        
        If pnITFValor > 0 Then
          If nITFConcepto = gConcITFCliente And (nITFOperacion = gITFCobroCargo Or nITFOperacion = gITFCobroCMACCargo) Then
            nITFDescuento = pnITFValor
          End If
        End If
                
        If sCopecod = "200901" Then
            If Not ValidaSaldoCuenta(sCuenta, nMonto - nITFDescuento) Then
                Set clsCap = Nothing
                Set rs = Nothing
                Set rsItf = Nothing
                MsgBox "Cuenta " & sCuenta & " NO POSEE SALDO SUFICIENTE PARA EXTORNAR ", vbInformation, "Aviso"
                TransValidaSaldos = False
                Exit Function
            End If
        End If
        rs.MoveNext
Wend

Set clsCap = Nothing
Set rs = Nothing
Set rsItf = Nothing

End Function


Public Function CapRetiroInteresPF(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nMonto As Double, ByVal nDias As Long, ByVal sGlosa As String, _
        ByVal nOperacion As CaptacOperacion, Optional sCuentaAho As String = "", _
        Optional nTipoCambio As Double, Optional sCodCMAC As String = "", _
        Optional sNomCmac As String, Optional sNomAge As String = "", Optional sLpt As String = "", _
        Optional sPersLavDinero As String = "", Optional ByVal psCodCMAC As String = "", _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional pnMovNro As Long = 0) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date, dAuxiliar As Date, dFechaProx As Date, sMenProx As String * 38
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long, nDiasAux As Long
Dim nExtracto As Long, nPlazo As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date, nFormaRetiro As Integer, dRenov As Date, dApertura As Date

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
nFormaRetiro = rsCta("nFormaRetiro")
dRenov = rsCta("dRenovacion")
dApertura = rsCta("dApertura")

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

sMenProx = ""

    Select Case nFormaRetiro
        Case 1
               dFechaProx = DateAdd("d", 30, dFecSis)
               sMenProx = "RET. MENSUAL PROX RETIRO:"
               
        Case 2
               dFechaProx = DateAdd("d", nPlazo, dRenov)
               sMenProx = "RET. FIN PLAZO PROX RENOV:"
               
        Case 3
               dFechaProx = DateAdd("d", nPlazo, dRenov)
               sMenProx = "RET. LIB PROX RENOV:"
    End Select
    
    sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
    

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

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
End If

''If sPersLavDinero <> "" Then
''    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
''End If
pnMovNro = nMovNro
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
        If CLng(Mid(sCuentaAho, 9, 1)) = gMonedaExtranjera Then
            nMonto = Round(nMonto / nTipoCambio, 2)
        Else
            nMonto = Round(nMonto * nTipoCambio, 2)
        End If
    End If
    
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMonto
    clsCap.ActualizaSaldoAnteriorAho sCuentaAho, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto
    clsCap.ActualizaAbonoCaptacion sCuentaAho, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Depósito Int. PF"
    
    clsCap.AgregaMovCap nMovNro, gAhoDepAboIntPF, sCuentaAho, nMonto, nSaldoDisp, nSaldoCnt
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
    
    clsCap.AgregaMovCapDet nMovNro, gAhoDepAboIntPF, sCuentaAho, gConcCapital, nMonto
    clsCap.UltimaActualizacionCuenta sCuentaAho, sMovNro
    
    clsCap.dbCmact.CommitTrans
    
    sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "DEPOSITO AHORRO"
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
    Set clsTit = Nothing
    
    Do
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuentaAho, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC
        
    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
    
Else
    
    clsCap.dbCmact.CommitTrans
    
End If


sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO PLAZO FIJO"


Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing

Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , Trim(nDias), , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , , , , , sMenProx
    
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, , nSaldoDisp, nSaldoCnt, CStr(nOperacion)
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, , nSaldoDisp, nSaldoCnt, CStr(nOperacion)
            End If
        Else
            If psItfOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, , nSaldoDisp, nSaldoCnt, CStr(nOperacion)
                
            Else
                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, , nSaldoDisp, nSaldoCnt, CStr(nOperacion)
            End If
        End If
    End If

    
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

bTrans = False

CapRetiroInteresPF = nSaldoCnt

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapRetiroInteresPF = 0
End Function

Public Function GetMontoRetiroIntPF(ByVal nCapital As Double, ByVal dUltMov As Date, _
        ByVal nTasa As Double, ByVal dFecha As Date, ByVal nIntAcum As Double) As Double
Dim nDiasTranscurridos As Long
Dim nIntGanado  As Double


nDiasTranscurridos = DateDiff("d", dUltMov, dFecha) - 1

If nDiasTranscurridos < -1 Then MsgBox "INFORME A SISTEMAS"

If nDiasTranscurridos < 0 Then nDiasTranscurridos = 0
 
nIntGanado = GetInteres(nCapital + nIntAcum, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
GetMontoRetiroIntPF = nIntAcum + nIntGanado
End Function

Public Function GetFechaUltimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
GetFechaUltimoRetiroIntPF = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
Set clsCap = Nothing
End Function

Public Function GetFechaPenultimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim clsCap As DCapMovimientos
Set clsCap = New DCapMovimientos
GetFechaPenultimoRetiroIntPF = clsCap.GetFechaPenultimoRetiroIntPF(sCuenta)
Set clsCap = Nothing
End Function

Public Function CapCancelacionPF(ByVal sCuenta As String, ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
            ByVal sGlosa As String, ByVal nCapital As Double, ByVal nIntGanado As Double, _
            Optional sNroDoc As String = "", Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional ByVal psCodCMAC As String = "", _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional SBENPERSLAVDINERO As String = "", _
            Optional bTransFERENCIA As Boolean = False, Optional sCodIF As String = "", Optional psCtaBanco As String = "", Optional pnMoneda As Moneda = gMonedaNacional, Optional pnMovNro As Long = 0) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nTasa As Double, nSaldoCanc As Double
Dim dUltMov As Date, dAuxiliar As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long, nDiasAux As Long
Dim nExtracto As Long, nPlazo As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String, sCodage As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
sCodage = Mid(sCuenta, 4, 2)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
dAuxiliar = rsCta("dAuxiliar")
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Verifica el estado de la cuenta por si acaso, no vaya a haber algun chistoso que haya
'hecho algo con la cuenta
If nEstado <> gCapEstActiva Then
    Select Case nEstado = gCapEstCancelada
        Case gCapEstCancelada, gCapEstAnulada
            MsgBox "Cuenta Cancelada o Anulada", vbInformation, "Aviso"
        Case gCapEstBloqTotal, gCapEstBloqRetiro
            MsgBox "Cuenta Bloqueada", vbInformation, "Aviso"
    End Select
    CapCancelacionPF = 0
    Exit Function
End If

Randomize
For i = 0 To Rnd(2000) * 1000
Next i

'Verifica si en alguna parte de la Caja se retiraron intereses antes de grabar
nSaldoCanc = GetSaldoCancelacion(sCuenta, dFecSis, sCodage)
If Round(nSaldoCanc, 2) <> Round(nSaldoDisp + nIntGanado, 2) Then
    MsgBox "Se retiraron intereses antes de efectuar la cancelación. " & Chr$(13) _
        & "Por favor consulte nuevamente el saldo de la cuenta", vbInformation, "Aviso"
    CapCancelacionPF = 0
    Exit Function
End If

If TieneChequesValorizacion(sCuenta) Then
    MsgBox "La Cuenta Tiene cheques en valorizacion", vbInformation, "Aviso"
    CapCancelacionPF = 0
    Exit Function
End If

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True

If sGlosa = "" Then
    sGlosa = "Cancelación PF Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & "Cancelación PF Cuenta = " & sCuenta
End If
clsCap.ActualizaEstadoCuenta sCuenta, gCapEstCancelada, dFecSis, sMovNro
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntGanado, 0, 0
If nIntGanado - nIntAcum > 0 Then
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvision, nIntGanado - nIntAcum
Else
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcRevProvision, Abs(nIntGanado - nIntAcum)
End If

If bTransFERENCIA = True Then

    Dim lnMontoCV As Double, nMonto As Double
    nMonto = (nSaldoDisp + nIntGanado)
    clsCap.AgregaMovTranferBanco nMovNro, sCodIF, psCtaBanco, pnMoneda


    If pnMoneda <> Mid(sCuenta, 9, 1) Then
        GetTipCambio dFecSis
        
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


If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
End If

clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
If nIntGanado < 0 Then
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcProvisionNeg, Abs(nIntGanado)
ElseIf nIntGanado > 0 Then
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntGanado
End If
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
''ALPA 20081010
''If sPersLavDinero <> "" Then
''    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero, SBENPERSLAVDINERO
''End If
pnMovNro = nMovNro
clsCap.dbCmact.CommitTrans
bTrans = False
CapCancelacionPF = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "CANCELACION PLAZO FIJO"
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    If sNroDoc <> "" Then
        ImprimeBoletaInteres sTipApe, "Capital", sCodOpe, Trim(nSaldoDisp), sNomTit, sCuenta, sNroDoc, 0, Trim(0), "Nota Cargo", nExtracto, "Interes Retirado", Trim(nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt, psCodCMAC
    Else
        ImprimeBoletaInteres sTipApe, "Capital", sCodOpe, Trim(nSaldoDisp), sNomTit, sCuenta, "", 0, Trim(0), "", nExtracto, "Interes Retirado", Trim(nIntGanado), , , , sNomAge, dFecSis, sCodUser, sLpt, psCodCMAC
    End If

    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            fgITFImprimeBoleta fgITFGetTitular(sCuenta), pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
        Else
            fgITFImprimeBoleta fgITFGetTitular(sCuenta), pnITFValor, sGlosa, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
        End If
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    CapCancelacionPF = 0
End Function

Public Sub EmiteDuplicadoCertPF(ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String)
Dim clsCap As DCapMovimientos
Dim clsMant As DCapMantenimiento
Dim nMovNro As Long
Dim bTrans As Boolean
Dim nSaldoDisp As Double, nSaldoCnt As Double
Dim rsCta As Recordset

bTrans = False
'Obtiene los datos de la Cuenta
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
On Error GoTo ErrGraba
Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = "Emisión Duplicado de Plazo Fijo Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & "Emisión Duplicado de Plazo Fijo Cuenta = " & sCuenta
End If
clsCap.AgregaMov sMovNro, gPFDupCert, sGlosa, gMovEstContabNoContable
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCap nMovNro, gPFDupCert, sCuenta, 0, nSaldoDisp, nSaldoCnt
clsCap.ActualizaDuplicadoCertPF sCuenta
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Function CapExtornoRetiroIntPF(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional ByVal sNomAge As String = "", Optional ByVal sLpt As String = "", Optional ByVal psCodCMAC As String, _
            Optional ByVal nITFValor As Double = 0, Optional ByVal nITFOperacion As CaptacOperacion, _
            Optional ByVal nITFConcepto As CaptacConcepto) As Double

Dim nMovNro As Long, nExtracto As Long, i As Long
Dim sMsgOpe As String
Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntGanado As Double
Dim nTasa As Double, nIntAcum As Double
Dim dAuxiliarAnt As Date, dUltMovAnt As Date
Dim dUltMov As Date
Dim nEstado As CaptacEstado
Dim bTrans As Boolean
Dim rsCta As Recordset
Dim nSaldoInac As Double, nMontoAho As Double
Dim bInactiva As Boolean
Dim sCuentaAho As String, sCodUser As String
Dim nDiasTranscurridos As Long
Dim dFecSis As Date
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim clsTit As NCapMantenimiento

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba

Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)




clsCap.dbCmact.BeginTrans
bTrans = True
'Actualiza y agrega operacion de extorno
If sGlosa = "" Then
    If nOperacion = gPFRetIntAboAho Then
        sGlosa = "Extorno Retiro Interes PF. Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
    Else
        sGlosa = "Extorno Retiro Interes PF. Cuenta = " & sCuenta
    End If
Else
    If nOperacion = gPFRetIntAboAho Then
        sGlosa = sGlosa & Chr$(13) & "Extorno Retiro Interes PF. Cuenta = " & sCuenta & ". Abono Cuenta Ahorros = " & sCuentaAho
    Else
        sGlosa = sGlosa & Chr$(13) & "Extorno Retiro Interes PF. Cuenta = " & sCuenta
    End If
End If
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
nMovNro = clsCap.GetnMovNro(sMovNro)

If nOperacion = gPFExtRetIntAboAho Then 'Si el extorno es con abono a cuenta de ahorros
    Set rsCta = clsCap.GetCapMov(nMovNroBus, gConcCapital, gCapAhorros)
    sCuentaAho = rsCta("cCtaCod")
    nMontoAho = rsCta("nMonto")
    'Obtiene los datos para el calculo
    Set clsMant = New DCapMantenimiento
    Set rsCta = clsMant.GetDatosCuentaAho(sCuentaAho)
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
    Randomize
    For i = 0 To Rnd(2000) * 1000
    Next i
    If Not ValidaSaldoCuenta(sCuentaAho, nMontoAho) Then
        clsCap.dbCmact.RollbackTrans
        Set clsCap = Nothing
        Err.Raise 1000, "CapExtornoRetiroIntPF", "Cuenta NO Posee Saldo Suficiente"
        CapExtornoRetiroIntPF = 0
        Exit Function
    End If
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt - nMonto
    clsCap.ActualizaSaldoAnteriorAho sCuentaAho, nSaldoDisp
    nSaldoDisp = nSaldoDisp - nMonto
    clsCap.ActualizaCargoCaptacion sCuentaAho, nMontoAho, nMontoAho, nIntGanado, dUltMov, sMovNro, True
    clsCap.AgregaMovCap nMovNro, nOperacion, sCuentaAho, nMontoAho, nSaldoDisp, nSaldoCnt
    clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuentaAho, gConcCapital, nMontoAho
    clsCap.UltimaActualizacionCuenta sCuentaAho, sMovNro
    sMsgOpe = "Ext.Ret.Int.PF Abo.Cta.Aho"
    
    sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "EXTORNO AHORROS"
    
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuentaAho))
    Set clsTit = Nothing
    Do
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMontoAho), sNomTit, sCuentaAho, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
    
Else
    sMsgOpe = "Ext. Retiro Interes PF"
End If

'Obtiene los datos para el cálculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
Set clsMant = Nothing
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
nEstado = rsCta("nPrdEstado")
dAuxiliarAnt = rsCta("dAuxiliarAnt")
dUltMovAnt = rsCta("dUltCierreAnt")
nExtracto = rsCta("nTransacc")
nIntAcum = rsCta("nIntAcum")
rsCta.Close
Set rsCta = Nothing
Set rsCta = clsCap.GetCapMov(nMovNroBus, gConcProvision, gCapPlazoFijo)
If Not (rsCta.EOF And rsCta.BOF) Then
    nIntGanado = rsCta("nMontoConcepto")
Else
    nIntGanado = 0
End If

clsCap.ActualizaExtornoRetiroInteresPF sCuenta, nMonto, nIntGanado, dUltMovAnt, sMovNro, dAuxiliarAnt
clsCap.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nMonto
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
CapExtornoRetiroIntPF = nSaldoCnt

sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO PLAZO FIJO"
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum - nIntGanado + nMonto, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus
    If nITFValor > 0 Then
        If nITFConcepto = gConcITFAsumido Then
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, nITFValor, "Ext. Retiro Int PF = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, nITFValor, "Ext. Retiro Int PF = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser
            End If
        Else
            If nITFOperacion = gITFCobroEfectivo Then
                fgITFImprimeBoleta sNomTit, nITFValor, "Ext. Retiro Int PF = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            Else
                fgITFImprimeBoleta sNomTit, nITFValor, "Ext. Retiro Int PF = " & sCuenta, dFecSis, sLpt, nExtracto, "EXT.IMP.TRANSAC.FINANC.", 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser
            End If
        End If
    End If
    
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "CapExtornoCancelacionAho", Err.Description
    CapExtornoRetiroIntPF = 0
End Function

Public Sub CapConsultaSaldosMovimiento(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)

Dim clsCap As DCapMovimientos
Dim nMovNro As Long
Dim bTrans As Boolean
Dim rsCta As Recordset
Dim sGlosa As String
bTrans = False
'Obtiene los datos de la Cuenta
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
    If bTrans Then clsCap.dbCmact.RollbackTrans
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
            Optional sLpt As String = "", Optional ByVal psCodCMAC As String) As Double

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
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)

Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")

dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nIntSaldo = rsCta("nIntSaldo")
nExtracto = rsCta("nTransacc") + 1
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
'If Not ValidaSaldoCuenta(sCuenta, nMonto, nOperacion) Then
'    clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise 1000, "CapExtornoAbonoCTS", "Cuenta NO Posee Saldo Suficiente"
'    CapExtornoAbonoCTS = 0
'    Exit Function
'End If
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
        clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo
        clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldRetiro = nSaldRetiro - (nMonto * nPorcDisp / 100)
    nSaldoDisp = nSaldoDisp - nMonto
    clsCap.ActualizaSaldoRetiroCTS sCuenta, nSaldRetiro, nIntSaldo
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Ext. Depósito Efectivo"
End If

clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
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
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    Else
        ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
    End If
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoAbonoCTS = 0
End Function

Public Sub CapOpeAhoCMACLlamada(ByVal sMovNro As String, ByVal nMoneda As Moneda, _
            ByVal sGlosa As String, ByVal nOperacion As CaptacOperacion, ByVal nExtracto As Integer, _
            ByVal sDescOperacion As String, ByVal nMonto As Double, ByVal sCuenta As String, _
            ByVal nSaldo As Double, ByVal sCodCMAC As String, ByVal sNomCmac As String, _
            ByVal sCliente As String, Optional nTpoDoc As TpoDoc = TpoDocOrdenPago, _
            Optional sDocumento As String = "NULL", Optional sCuentaAho As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional nComision As Double = 0, Optional nMonedaComision As Moneda = gMonedaNacional, Optional psPersLavDinero As String = "", Optional pnITFValorCom As Currency = 0, Optional psPersBenLavDinero As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgOpe As String, sCodUser As String, sCodage As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long
Dim oCont As NContFunciones
Dim nTempo As Double

Dim sMovNroComITF As String
Dim nMovNroComITF As String

Dim nBan As Boolean

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
sCodage = Mid(sMovNro, 18, 2)
bTrans = False

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
clsCap.bTransaccion = True

bTrans = True
If sGlosa = "" Then
    sGlosa = sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & Chr$(13) & sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
End If
clsCap.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovCMAC nMovNro, sCodCMAC, nMoneda, nOperacion, nMonto, sDocumento, sCuenta, sCliente
If Len(Trim(sDocumento)) > 0 And nTpoDoc = TpoDocOrdenPago And nOperacion = gCMACOTAhoRetOP Then
    clsCap.AgregaMovDoc nMovNro, nTpoDoc, sDocumento, dFecSis
End If
'xxx eliminado una transacc clsCap.dbCmact.CommitTrans
'Fin transaccion

'bTrans = False

If nComision > 0 Then
    Dim sMovNroCom As String
    Dim nMovNroCom As Long
    Set oCont = New NContFunciones
    
    'xxx Agregado
    sMovNroCom = oCont.IncrementaMovNro(sMovNro)
        
    'xxx eliminado sMovNroCom = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)
    Set oCont = Nothing
    
    'Inicio Transaccion
    'xxx eliminado clsCap.dbCmact.BeginTrans
    
    clsCap.AgregaMov sMovNroCom, gOtrOpeComisionCMACLlam, sGlosa
    nMovNroCom = clsCap.GetnMovNro(sMovNroCom)
    clsCap.AgregaMovOpeVarias nMovNroCom, nComision, sCuenta, sCodCMAC, nMonedaComision
    clsCap.AgregaMovRef nMovNro, nMovNroCom
    
    If pnITFValorCom > 0 Then
        sMovNroComITF = Left(nMovNroCom, 2) & "2" & Right(nMovNroCom, 4)
        clsCap.AgregaMov sMovNroComITF, gITFCobroEfectivo, sGlosa
        nMovNroComITF = clsCap.GetnMovNro(sMovNroComITF)
        clsCap.AgregaMovOpeVarias nMovNroComITF, pnITFValorCom, sCuenta, sCodCMAC, nMonedaComision
        clsCap.AgregaMovRef nMovNroCom, nMovNroComITF
    End If

End If

If sCuentaAho <> "" And nOperacion <> gCMACOTAhoDepChq Then
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Set oCont = New NContFunciones
    
    'xxx agregado
    If Len(Trim(sMovNroCom)) > 0 Then
        sMovNroReg = oCont.IncrementaMovNro(sMovNroCom)
    Else
        sMovNroReg = oCont.IncrementaMovNro(sMovNro)
    End If
    
    Set oCont = Nothing
    Select Case nOperacion
        Case gCMACOTAhoDepEfec
            nTempo = CapAbonoCuentaAho(sCuentaAho, nMonto, gAhoDepRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , , clsCap.dbCmact, , , , , , , psPersBenLavDinero)
            If nTempo = 0 Then
              ' If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
                clsCap.dbCmact.RollbackTrans
                Set clsCap = Nothing
                Exit Sub
            End If
        Case gCMACOTAhoRetEfec, gCMACOTAhoRetOP
            nTempo = CapCargoCuentaAho(sCuentaAho, nMonto, gAhoRetRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , clsCap.dbCmact, , , , , , , , psPersBenLavDinero)
            If nTempo = 0 Then
                'If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
                clsCap.dbCmact.RollbackTrans
                Set clsCap = Nothing
                Exit Sub
            End If
    End Select
    nMovNroReg = clsCap.GetnMovNro(sMovNroReg)
    clsCap.AgregaMovRef nMovNro, nMovNroReg
End If

clsCap.dbCmact.CommitTrans
clsCap.bTransaccion = False

'Imprime Documentos
MsgBox "Se imprimirán las boletas de regularización", vbInformation, "Aviso"
Dim vVez As Integer

vVez = 1
nBan = True
Do While nBan = True

    If sDocumento = "NULL" Then
        ImprimeBoletaCMAC "OPERACION CMAC LLAMADA", sDescOperacion, Trim(nMonto), sCliente, sCuenta, "", nSaldo, nMoneda, "", nExtracto, nSaldo, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, False, sCodage
    Else
        ImprimeBoletaCMAC "OPERACION CMAC LLAMADA", sDescOperacion & " N° " & sDocumento, Trim(nMonto), sCliente, sCuenta, "", nSaldo, nMoneda, "", nExtracto, nSaldo, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, False, sCodage
    End If
    
    If vVez = 1 Then
        ImprimeBoletaCMAC "OPERACION CMAC LLAMADA", "Comision Ope Llam", Trim(nComision), sCliente, sCuenta, "", 0, nMonedaComision, "", 0, 0, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, False, sCodage
    Else
        If MsgBox("Desea re-imprimir Boleta de Comisión ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            ImprimeBoletaCMAC "OPERACION CMAC LLAMADA", "Comision Ope Llam", Trim(nComision), sCliente, sCuenta, "", 0, nMonedaComision, "", 0, 0, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, False, sCodage
        End If
    
    End If
    vVez = vVez + 1

    nBan = MsgBox("Desea re-imprimir Boletas de regularizacion ?", vbQuestion + vbYesNo, "Aviso") = vbYes

Loop
Set clsCap = Nothing

Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    
End Sub

Public Sub CapExtornoOpeAhoCMACLlamada(ByVal nMovNroBus As Long, ByVal sMovNro As String, _
            ByVal nOperacion As CaptacOperacion, ByVal sGlosa As String, ByVal sCuenta As String, _
            ByVal sDescOperacion As String, ByVal nMonto As Double, _
            Optional sDocumento As String = "NULL", Optional sNomAge As String, _
            Optional sLpt As String = "", Optional ByVal psCodCMAC As String)

Dim clsCap As DCapMovimientos
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long
Dim RSTEMP As Recordset

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Inicia la transaccion
Set clsCap = New DCapMovimientos

Dim clsMov As DMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New DMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(nMovNroBus)
sCodUserBus = Right(sMovNroBus, 4)



'On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True


sDocumento = IIf(sDocumento = "NULL", "''", sDocumento)

Set RSTEMP = clsCap.MovAsociados(nMovNroBus)
Select Case nOperacion
 Case gCMACOTAhoExtDepEfec, gCMACOTAhoExtDepChq, "137000"
    ActualizaDepCMACLLAMADA nMovNroBus, gAhoDepRegCMACLlam, sDocumento, sMovNro
 Case gCMACOTAhoExtRetEfec, gCMACOTAhoExtRetOP
    ActualizaRetCMACLLAMADA nMovNroBus, gAhoRetRegCMACLlam, sDocumento, sMovNro
End Select

clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
While Not RSTEMP.EOF
    clsCap.ActualizaEstadoMov RSTEMP!nMovNroRef, gMovFlagExtornado
    RSTEMP.MoveNext
Wend
If RSTEMP.State = 1 Then RSTEMP.Close
Set RSTEMP = Nothing


clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
clsCap.dbCmact.CommitTrans
If sDocumento = "" Then
    ImprimeBoleta "EXT.OPE CMAC LLAMADA", sDescOperacion, Trim(nOperacion), Trim(nMonto), "", sCuenta, "", 0, "", "", 0, 0, True, False, , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
Else
    ImprimeBoleta "EXT.OPE CMAC LLAMADA", sDescOperacion & " N° " & sDocumento, Trim(nOperacion), Trim(nMonto), "", sCuenta, "", 0, "", "", 0, 0, True, False, , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, gsCodAge, , , CStr(nOperacion), sCodUserBus, True
End If
'Exit Sub
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'   Set clsCap = Nothing
 '   Err.Raise Err.Number, "", Err.Description
End Sub
Public Sub ActualizaRetCMACLLAMADA(ByVal nMovNro As Long, ByVal nTipoOpe As CaptacOperacion, ByVal cDocumento As String, ByVal sMovNro As String)
 Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim nDiasTranscurridos As Long
Dim dFecSis As Date


dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
bTrans = False
'Obtiene los datos para el calculo

Set clsMant = New DCapMantenimiento
Set clsCap = New DCapMovimientos
Set RSTEMP = clsMant.GetDatosOpeLLamada(nMovNro, CStr(nTipoOpe))
    sCuenta = RSTEMP("cctacod")
    nMonto = RSTEMP("nmonto")

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

On Error GoTo ErrGraba
clsCap.dbCmact.BeginTrans
bTrans = True

nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp


clsCap.ActualizaAbonoCaptacion sCuenta, nMonto + nITFMonto, nMonto + nITFMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "''" Then
    If nTipoDoc = TpoDocOrdenPago Then
         clsCap.EliminaOrdenPago sCuenta, sNroDoc
    End If
End If

clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
'CapExtornoCargoAho = nSaldoCnt



Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "ACTUALIZA EXTORNO RETIRO CMAC LLAMADA", Err.Description
    'CapExtornoCargoAho = 0

End Sub


Public Sub ActualizaDepCMACLLAMADA(ByVal nMovNro As Long, ByVal nTipoOpe As CaptacOperacion, ByVal cDocumento As String, ByVal sMovNro As String)
Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset, RSTEMP As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long
Dim dFecSis As Date
Dim sCuenta As String, nMonto As Double

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))

bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set clsCap = New DCapMovimientos

Set RSTEMP = clsMant.GetDatosOpeLLamada(nMovNro, CStr(nTipoOpe))


    sCuenta = RSTEMP("cctacod")
    nMonto = RSTEMP("nmonto")


Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")

rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion

On Error GoTo ErrGraba
bTrans = True

If sNroDoc = "" Then
    If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
        Set clsCap = Nothing
        MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
        'ActualizaDepCMACLLAMADA = 0
        Exit Sub
    End If
End If
'sGlosa = sGlosa & ". Extorno Abono Cuenta = " & sCuenta
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
        'clsCap.ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, smovnro, gsChqEstExtornado
      
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    clsCap.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
End If
clsCap.UltimaActualizacionCuenta sCuenta, sMovNro
clsCap.dbCmact.CommitTrans
bTrans = False


Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"


End Sub
'*******************************

Public Sub CapCargoFondoFijo(ByVal sNroDoc As String, ByVal nMonto As Double, ByVal nMovNroRef As Long, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal dFecDoc As Date, _
            ByVal sGlosa As String, Optional sNomAge As String = "", Optional sLpt As String = "", Optional pnMoneda As Moneda = Moneda.gMonedaNacional)

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
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagVigente  'DeExtorno

nMovNro = clsCap.GetnMovNro(sMovNro)
clsCap.AgregaMovDoc nMovNro, TpoDocOrdenPago, sNroDoc, dFecDoc
clsCap.AgregaMovRef nMovNro, nMovNroRef
clsCap.AgregaMovOpeVarias nMovNro, nMonto, sNroDoc, sGlosa, pnMoneda, nOperacion

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
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function CtaConFirmas(psCtaCod As String) As Boolean
    Dim oCapMov As DCapMovimientos
    Set oCapMov = New DCapMovimientos
    
    CtaConFirmas = oCapMov.CtaConFirmas(psCtaCod)
    
    Set oCapMov = Nothing
End Function

Public Function OtrasOperaciones(psMovNro As String, pnOperacion As CaptacOperacion, pnMovImporte As Currency, psNroDoc As String, psReferencia As String, pnMoneda As Moneda, psPersCod As String, Optional pnMovNroRef As Long = 0) As Long
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim lnMovNro As Long
    
    clsCapMov.AgregaMov psMovNro, pnOperacion, psReferencia, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = clsCapMov.GetnMovNro(psMovNro)
    clsCapMov.AgregaMovOpeVarias lnMovNro, pnMovImporte, psNroDoc, psReferencia, pnMoneda, pnOperacion
    clsCapMov.AgregaMovGasto lnMovNro, psPersCod, ""
    
    If pnMovNroRef <> 0 Then
        clsCapMov.AgregaMovRef pnMovNroRef, lnMovNro
    End If
    
    Set clsCapMov = Nothing
    OtrasOperaciones = lnMovNro
End Function

Public Function GetTarifa(psOpeCod As CaptacOperacion, pnMoneda As Moneda) As Currency
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    
    GetTarifa = clsCapMov.GetTarifa(psOpeCod, pnMoneda)
    
    Set clsCapMov = Nothing
End Function

Public Function GetOtrasOperaciones(ByVal pdFecha As Date, Optional cUser As String) As ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = clsCapMov.GetOtrasOperaciones(pdFecha, cUser)
    
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

Public Function GetCreditosPendientes(ByVal sPersona As String, pgdFecSis As Date, Optional ByVal pbBusqCred As Boolean = True, _
        Optional pbBusqPrend As Boolean = True, Optional pbBusqJud As Boolean = True) As ADODB.Recordset
    Dim rsCred As ADODB.Recordset
    Set rsCred = New ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
    Dim nCuentasCre As Integer
    Dim nCuentasPrd As Integer
    Dim nCuentasJud As Integer
    
    nCuentasCre = 0
    nCuentasPrd = 0
    nCuentasJud = 0
    
    GeneraRecordsetCred rsCred
    rsCred.Open
    
    ReDim Preserve CuentasPersonaCre(0)
    If pbBusqCred = True Then
        Set rsCta = clsCapMov.CargaCuentas(sPersona, "CRE")
        
        If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasCre = nCuentasCre + 1
                ReDim Preserve CuentasPersonaCre(nCuentasCre)
                CuentasPersonaCre(nCuentasCre) = rsCta("cCtaCod")
                rsCta.MoveNext
            Loop
        End If
    End If
    
    ReDim Preserve CuentasPersonaPrd(0)
    If pbBusqPrend = True Then
        Set rsCta = clsCapMov.CargaCuentas(sPersona, "PRD")
        If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasPrd = nCuentasPrd + 1
                ReDim Preserve CuentasPersonaPrd(nCuentasPrd)
                CuentasPersonaPrd(nCuentasPrd) = rsCta("cCtaCod")
                rsCta.MoveNext
            Loop
        End If
    End If
    
    ReDim Preserve CuentasPersonaJud(0)
    If pbBusqJud = True Then
        Set rsCta = clsCapMov.CargaCuentas(sPersona, "JUD")
        If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasJud = nCuentasJud + 1
                ReDim Preserve CuentasPersonaJud(nCuentasJud)
                CuentasPersonaJud(nCuentasJud) = rsCta("cCtaCod")
                rsCta.MoveNext
            Loop
        End If
    End If
    CreditosPendientes sPersona, rsCred, pgdFecSis
    
    Set GetCreditosPendientes = rsCred
End Function

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

Public Function CreditosPendientes(ByVal sPersona As String, ByRef rsCred As ADODB.Recordset, pgdFecSis As Date) As Boolean
    Dim rsCta As ADODB.Recordset
    Dim nAgencias As Long, nLen As Long, nCuentasCre As Long, nCuentasPrd As Long, nCuentasJud As Long
    Dim sBuffer As String, sAgencia As String, sTipo As String
    Dim sSimbMoneda As String, sCuentas As String, sVencidos As String
    Dim i As Integer, j As Integer, nDiasAtraso As Integer
    Dim nDeudaPrend As Double, nMinimoPagar As Double
    Dim bVencidos As Boolean
    Dim nValor1 As Double, nValor2 As Double
    Dim clsCapDef As nCapDefinicion
    Set clsCapDef = New nCapDefinicion
    
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
     
    sAgencia = ""
    sBuffer = ""
    nCuentasCre = 0
        
    nValor1 = clsCapDef.GetCapParametro(2027)
    nValor2 = clsCapDef.GetCapParametro(2028)
    
    If nValor1 = 1 Then
        bVencidos = True
        nDiasAtraso = Int(nValor2)
    Else
        bVencidos = False
    End If
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    
    For i = 1 To UBound(CuentasPersonaCre, 1)
        sCuentas = "("
        sAgencia = Left(CuentasPersonaCre(i), 2)
        For j = i To UBound(CuentasPersonaCre, 1)
            sBuffer = Left(CuentasPersonaCre(j), 2)
            If sBuffer = sAgencia Then
                sCuentas = sCuentas & "'" & CuentasPersonaCre(j) & "',"
            Else
                Exit For
            End If
        Next j
        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
        sCuentas = sCuentas & ")"
        i = j - 1
        
        Set rsCta = clsCapMov.GetDatosCrePendiente("CRE", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasCre = nCuentasCre + 1
                rsCred.AddNew
                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
                rsCred("Estado") = IIf(rsCta("cRefinan") = 1, "RF", "")
                rsCred("Vencimiento") = CDate(rsCta("dFecVenc"))
                rsCred("CuoVenc") = Trim(rsCta("nNroCuo"))
                rsCred("Cuota") = rsCta("nCuota")
                rsCred("Mora") = rsCta("nMora")
                rsCred("Gastos") = rsCta("nGastos")
                rsCred("MinPagar") = rsCta("nCuota") + rsCta("nGastos") + rsCta("nMora")
                rsCred.Update
                rsCta.MoveNext
            Loop
        End If
        rsCta.Close
    Next i
    
    'Verifica los Créditos Prendarios
    sAgencia = ""
    sBuffer = ""
    nCuentasPrd = 0
    For i = 1 To UBound(CuentasPersonaPrd, 1)
        sCuentas = "("
        sAgencia = Left(CuentasPersonaPrd(i), 2)
        For j = i To UBound(CuentasPersonaPrd, 1)
            sBuffer = Left(CuentasPersonaPrd(j), 2)
            If sBuffer = sAgencia Then
                sCuentas = sCuentas & "'" & CuentasPersonaPrd(j) & "',"
            Else
                Exit For
            End If
        Next j
        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
        sCuentas = sCuentas & ")"
        i = j - 1

    Set rsCta = clsCapMov.GetDatosCrePendiente("CRE", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
    If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasPrd = nCuentasPrd + 1
                nDeudaPrend = CalculaDeudaPrendario(rsCta("nSaldoCap"), rsCta("dFecVenc"), rsCta("nValTasac"), rsCta("nTasaIntVenc"), _
                                    rsCta("nCostCusto"), rsCta("nTasaImpuesto"), rsCta("nPRDEstado"), rsCta("nCostoPrepRemate"), pgdFecSis)
                nMinimoPagar = CalculaMinimoPagar(rsCta("nTasaInt"), rsCta("nPlazo"), rsCta("nValTasac"), rsCta("nTasaCustodia"), _
                                    rsCta("nSaldoCap"), rsCta("nPRDEstado"), rsCta("dFecVenc"), rsCta("nTasaIntVenc"), rsCta("nTasaCustodia"), _
                                    rsCta("nTasaImpuesto"), rsCta("nCostoPrepRemate"), pgdFecSis)
                rsCred.AddNew
                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
                rsCred("Estado") = IIf(rsCta("nPrdEstado") = ColocEstado.gColPEstPRema, "RE", "")
                rsCred("Vencimiento") = rsCta("dFecVenc")
                rsCred("CuoVenc") = 1
                rsCred("Cuota") = rsCta("nSaldoCap")
                rsCred("Mora") = nDeudaPrend - rsCta("nSaldoCap")
                rsCred("Gastos") = 0
                rsCred("MinPagar") = nMinimoPagar
                rsCred.Update
                rsCta.MoveNext
            Loop
        End If
        rsCta.Close
        
    Next i
    
    nCuentasJud = 0
    
    If UBound(CuentasPersonaJud, 1) > 0 Then
        sCuentas = "("
        For i = 1 To UBound(CuentasPersonaJud, 1)
            sCuentas = sCuentas & "'" & CuentasPersonaJud(i) & "',"
        Next i
        sCuentas = Left(sCuentas, Len(sCuentas) - 1)
        sCuentas = sCuentas & ")"
        'Verificamos los Creditos Judiciales concentrados en la agencia principal
        
        Set rsCta = clsCapMov.GetDatosCrePendiente("PRD", bVencidos, sCuentas, nDiasAtraso, pgdFecSis)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Do While Not rsCta.EOF
                nCuentasJud = nCuentasJud + 1
                rsCred.AddNew
                rsCred("Cuenta") = Trim(rsCta("cCtaCod"))
                rsCred("Estado") = "JU"
                rsCred("Vencimiento") = Null
                rsCred("CuoVenc") = 0
                rsCred("Cuota") = 0
                rsCred("Mora") = 0
                rsCred("Gastos") = 0
                rsCred("MinPagar") = 0
                rsCred.Update
                rsCta.MoveNext
            Loop
        End If
        rsCta.Close
    End If
    
    If (nCuentasCre + nCuentasPrd + nCuentasJud) = 0 Then
        CreditosPendientes = False
    Else
        CreditosPendientes = True
    End If
End Function

Public Function BuscaCreditosPendientesPago(ByVal psCuenta As String) As Boolean
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos

    BuscaCreditosPendientesPago = clsCapMov.BuscaCreditosPendientesPago(psCuenta)
    
    Set clsCapMov = Nothing
End Function

Public Function GetFirma(psPersCod As String, Optional ByVal psCodAge As String = "") As ADODB.Recordset
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
        
    Set GetFirma = clsCapMov.GetFirma(psPersCod, psCodAge)

    Set clsCapMov = Nothing
End Function

Public Function GetVerificaRetiro(ByVal sCuenta As String) As Boolean
    Dim clsCapMov As DCapMovimientos
    Set clsCapMov = New DCapMovimientos
        
     GetVerificaRetiro = clsCapMov.GetVerificaRetiro(sCuenta)
    Set clsCapMov = Nothing


End Function


Public Function CalculaDeudaPrendario(pnSaldoCapital As Double, _
    pdFecVencimiento As Date, pnValortasacion As Double, _
    pnTasaInteresVencido As Double, pnCostoCustodiaMoratoria As Double, _
    pnTasaImpuesto As Double, pnEstado As ColocEstado, _
    pnCostoPreparacionRemate As Double, pgdFecSis As Date, Optional pdFecParaDeuda As String) As Double
    Dim vDiasAtra As Double
    Dim vDeud As Double
    Dim vInteMora As Double
    Dim vImpu As Double
    Dim vCostCustMora As Double
    Dim clsCalculos As NColPCalculos
    Set clsCalculos = New NColPCalculos
    
    If Len(pdFecParaDeuda) <= 0 Then pdFecParaDeuda = pgdFecSis
    vDiasAtra = IIf(DateDiff("d", pdFecVencimiento, pdFecParaDeuda) <= 0, 0, DateDiff("d", pdFecVencimiento, pdFecParaDeuda))
    If vDiasAtra = 0 Then
       vInteMora = 0
       vCostCustMora = 0
       vImpu = 0
    Else
       vInteMora = clsCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresVencido, vDiasAtra)
       vInteMora = Round(vInteMora, 2)
       vCostCustMora = clsCalculos.nCalculaCostoCustodiaMoratorio(pnValortasacion, pnCostoCustodiaMoratoria, vDiasAtra)
       vCostCustMora = Round(vCostCustMora, 2)
       vImpu = (vInteMora + vCostCustMora) * pnTasaImpuesto
       vImpu = Round(vImpu, 2)
    End If
    vDeud = Round(pnSaldoCapital, 2) + vInteMora + vCostCustMora + vImpu
    If pnEstado = ColocEstado.gColPEstPRema Or pnEstado = ColocEstado.gColPEstAdjud Then      ' Si esta en via de Remate
       vDeud = vDeud + Round((pnCostoPreparacionRemate * pnValortasacion), 2)
    End If
    CalculaDeudaPrendario = vDeud
    Set clsCalculos = Nothing
End Function

Public Function CalculaMinimoPagar(pnTasaInteres As Double, pnPlazo As Integer, pnValortasacion As Double, _
                                    pnTasaCustodia As Double, pnSaldoCapital As Double, pnEstado As ColocEstado, _
                                    pdFecVencimiento As Date, pnTasaInteresVencido As Double, _
                                    pnCostoCustodiaMoratoria As Double, pnTasaImpuesto As Double, _
                                    pnCostoPreparacionRemate As Double, pgdFecSis As Date, Optional pdFecParaDeuda As String) As Double
    ' **************************************
    ' ** Calculo del Monto Minimo a Pagar **
    ' **************************************
    '    Dim NumeroDias As Single
    Dim vFactor As Single
    Dim vdiasAtraso As Double
    Dim vInteres As Double
    Dim vCostoCustodia As Double
    Dim vImpuesto As Double
    Dim vInteresVencido As Double
    Dim vCostoCustodiaVencido As Double
    Dim vCostoPreparacionRemate As Double
    Dim vMontoMinimo As Double
    
    Dim clsCalculos As NColPCalculos
    Set clsCalculos = New NColPCalculos
    
    vInteres = 0
    vFactor = 1 - 1 / (1 + pnTasaInteres) ^ (pnPlazo / 360)
    If Len(pdFecParaDeuda) <= 0 Then pdFecParaDeuda = pgdFecSis
    vdiasAtraso = IIf(DateDiff("d", pdFecVencimiento, pdFecParaDeuda) <= 0, 0, DateDiff("d", pdFecVencimiento, pdFecParaDeuda))
    vCostoCustodia = clsCalculos.nCalculaCostoCustodia(pnValortasacion, pnTasaCustodia, pnPlazo)
    vCostoCustodia = Round(vCostoCustodia, 2)
    vInteres = pnSaldoCapital * vFactor
    vInteres = Round(vInteres, 2)
    If vdiasAtraso = 0 Then
       vInteresVencido = 0
       vCostoCustodiaVencido = 0
    Else
       vInteresVencido = clsCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresVencido, vdiasAtraso)
       vInteresVencido = Round(vInteresVencido, 2)
       vCostoCustodiaVencido = clsCalculos.nCalculaCostoCustodiaMoratorio(pnValortasacion, pnCostoCustodiaMoratoria, vdiasAtraso)
       vCostoCustodiaVencido = Round(vCostoCustodiaVencido, 2)
    End If
    vImpuesto = (vInteresVencido + vInteres + vCostoCustodia + pnCostoCustodiaMoratoria) * pnTasaImpuesto
    vImpuesto = Round(vImpuesto, 2)
    
    If pnEstado = ColocEstado.gColPEstPRema Or pnEstado = ColocEstado.gColPEstAdjud Then  ' Si esta en via de Remate
       vCostoPreparacionRemate = Round((pnCostoPreparacionRemate * pnValortasacion), 2)
    End If
    
    vMontoMinimo = vInteresVencido + vCostoCustodiaVencido + vInteres + vCostoCustodia + vImpuesto + vCostoPreparacionRemate
    vMontoMinimo = Round(vMontoMinimo, 2)
    
    CalculaMinimoPagar = vMontoMinimo
    Set clsCalculos = Nothing
End Function

Public Function TieneMovDespuesApertura(ByVal sCuenta As String) As Boolean
Dim oCapMov As DCapMovimientos
Set oCapMov = New DCapMovimientos
TieneMovDespuesApertura = oCapMov.TieneMovDespuesApertura(sCuenta)
Set oCapMov = Nothing
End Function

Public Function ServGiroExtornoApertura(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
        Optional sNomAge As String = "", Optional sLpt As String = "")

Dim clsCap As DCapMovimientos
Dim sMsgProd As String, sCodUser As String
Dim nProd As Producto
Dim nMovNro As Long
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)

sMsgProd = "EXTORNO GIRO"

Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
clsCap.ActualizaEstadoGiro sCuenta, gCapEstAnulada, dFecSis
clsCap.dbCmact.CommitTrans

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = "Ext. Apert. Giro"
sTipApe = sMsgProd
Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", 0, 0, "", 1, 0, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
End Function

Public Sub ServGiroExtornoCancelacion(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psCodCMAC As String)

Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
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

'Inicia la transaccion
On Error GoTo ErrGraba
Set clsCap = New DCapMovimientos
clsCap.dbCmact.BeginTrans
bTrans = True
If sGlosa = "" Then
    sGlosa = "Extorno Cancelacion Giro = " & sCuenta
Else
    sGlosa = sGlosa & ". Extorno Cancelación Giro = " & sCuenta
End If
clsCap.ActualizaEstadoGiro sCuenta, gCapEstActiva, dFecSis
sMsgOpe = "Ext. Canc. Efectivo"

clsCap.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
clsCap.AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno

clsCap.dbCmact.CommitTrans
bTrans = False
Set clsCap = Nothing

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe

sTipApe = "EXTORNO GIRO"

Dim clsTit As NCapMantenimiento
Set clsTit = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
Set clsTit = Nothing
Do
    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntAcum, "", 1, nSaldoCnt, , , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
Set clsTit = Nothing
Exit Sub
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function ValidaEstadoCuenta(ByVal sCuenta As String, Optional pbParaRetiro As Boolean = True) As Boolean
Dim clsMant As NCapMantenimiento
Set clsMant = New NCapMantenimiento
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Dim nEstado As CaptacEstado

Set rsCta = clsMant.GetDatosCuenta(sCuenta)

nEstado = rsCta("nPrdEstado")

rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing

If pbParaRetiro Then
    If nEstado = gCapEstActiva Then
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
Else
    If nEstado = gCapEstActiva Or nEstado = gCapEstBloqRetiro Or nEstado = gCapEstBloqTotal Then
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
End If
End Function

Public Function ValidaChequesEnValorizacion(ByVal sCuenta As String) As Boolean
Dim clsMant As NCapMantenimiento
Set clsMant = New NCapMantenimiento
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Dim nNum As Integer
Dim nMonto As Integer

Set rsCta = clsMant.GetChequesEnValorizados(sCuenta)

nNum = rsCta.Fields(0)
nMonto = rsCta.Fields(1)

rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing

If nNum = 0 Then
    ValidaChequesEnValorizacion = False
Else
    ValidaChequesEnValorizacion = True
End If

End Function

Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double, Optional nOperacion As CaptacOperacion) As Boolean
Dim nMoneda As Moneda
Dim nProducto As Producto
Dim clsMant As NCapMantenimiento
Dim clsPar As nCapDefinicion
Dim rsCta As Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Dim bOrdPag As Boolean
Dim nPersoneria As PersPersoneria
Dim nEstado As CaptacEstado

Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
nProducto = CLng(Mid(sCuenta, 6, 3))
bOrdPag = False
If nProducto = gCapAhorros Then
    nSaldo = rsCta("nSaldoDisp")
    bOrdPag = rsCta("bOrdPag")
ElseIf nProducto = gCapCTS Then
    If nOperacion = gCTSDepChq Or nOperacion = gCTSDepEfec Or nOperacion = gCTSDepPlanRRHH Or nOperacion = gCTSDepTransf Or nOperacion = "220303" Then
        nSaldo = rsCta("nSaldo")
    Else
        nSaldo = rsCta("nSaldRetiro") '+ rsCta("nIntsaldo")
    End If
End If

nPersoneria = rsCta("nPersoneria")
nEstado = rsCta("nPrdEstado")

rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))

If nEstado <> gCapEstActiva Then
    ValidaSaldoCuenta = False
Else
    If CLng(Mid(sCuenta, 6, 3)) = gCapAhorros Then
        Set clsPar = New nCapDefinicion
        nMontoMinimo = clsPar.GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
        Set clsPar = Nothing
    Else
        nMontoMinimo = 0
    End If
    If nSaldo - nMontoMinimo - nMonto >= 0 Then
        ValidaSaldoCuenta = True
    Else
        ValidaSaldoCuenta = False
    End If
End If
End Function

Public Function GetChequesEnValorizados(sCuenta) As Recordset
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
Set GetChequesEnValorizados = oCap.GetChequesEnValorizacion(sCuenta)
Set oCap = Nothing
End Function

Private Function ObtieneCargosPermitido(ByVal nOperacion As CaptacOperacion, ByVal pnMoneda As Integer, ByVal nMonto As Double) As ADODB.Recordset
                                

Dim lrst As ADODB.Recordset
Dim lsConn As DConecta

Dim lsql As String


Set lsConn = New DConecta

lsConn.AbreConexion


lsql = "select Cg.*,Ct.cRHCargoDescripcion from  CapAutorizacionRango Cg left join rhcargostabla Ct on Cg.cRhCargoCod =  Ct.cRhCargoCod" & _
       " where Cg.cOpecod =" & nOperacion & " and  ((" & nMonto & "<= Cg.nMontoIniSol) or (" & nMonto & " between Cg.nMontoIniSol and  Cg.nMontoFinSol ) " & _
       "or 1<>" & pnMoneda & ") and  ((" & nMonto & " <= Cg.nMontoIniDol) or (" & nMonto & "  between Cg.nMontoIniDol and  Cg.nMontoFinDol) or 2<>" & pnMoneda & ")" & _
       "order by Cg.cOpeCod, Cg.nMontoIniSol"

Set lrst = lsConn.CargaRecordSet(lsql)
Set ObtieneCargosPermitido = lrst

End Function

Private Function OpePorAutorizar(ByVal psCuenta As String, ByVal nOperacion As CaptacOperacion) As Boolean
    
Dim lsOperacion As CaptacOperacion
Dim lrst As ADODB.Recordset
Dim lsConn As DConecta

Dim lsql As String
Dim lnNumReg As Integer


Set lsConn = New DConecta

lsConn.AbreConexion


lsql = "Select count(*)%3 as NumReg from MovCapAutorizacionOpe mcao where mcao.cCtaCod = '" & psCuenta & _
       "' and  mcao.cOpeCodOri = '" & nOperacion & "'"
       
'       and  mcao.cOpeCod = '" & gAhoSolAprobOperacion & _
'       "' or  mcao.cOpeCod = '" & gAhoAutorizaOperacion & "' or  mcao.cOpeCod = '" & gAhoRechazaOperacion & _
'       "' or  mcao.cOpeCod = '" & gAhoAnulaOperacion & "'"

Set lrst = lsConn.CargaRecordSet(lsql)

lnNumReg = lrst("NumReg")

Select Case lnNumReg
    Case 0
        OpePorAutorizar = False
    Case 1
        OpePorAutorizar = True
    Case 2
        OpePorAutorizar = True
        'MsgBox "Ud. ha sido Autorizado para Relizar esta Operación", vbInformation, "Atención"
    Case 3
        OpePorAutorizar = False
End Select

End Function

Public Sub fgITFImprimeBoletaTrans(ByVal sNomCli As String, ByVal nMonITF As Double, ByVal psDesOpeOrigen As String, _
        ByVal pdFecha As Date, ByVal sLpt As String, Optional nNumTran As Long = 1, _
        Optional sTitBol As String = "IMP. TRANSAC. FINANCIERAS", Optional pnTipoPago As Integer = 1, _
        Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", Optional sGlosa As String, _
        Optional pCancelaImp As Boolean = True, Optional sNomAge As String = "", Optional sCodUser As String = "" _
        , Optional bSaldos As Boolean = False, Optional nSaldoDisp As Double, Optional nSaldoCont As Double, Optional nVeces As Integer = 0)

If pCancelaImp Then
    Exit Sub 'cancela impresion de itf
End If

Dim psTit As String
Dim sMonto As String
Dim psDesOpe As String
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lnTope As Integer
Dim lsNroExt As String
Dim lnCliAux As Integer
Dim lsNomAge As String
Dim lnSaldoDisp As Double
Dim lnSaldoCnt As Double


Dim lnNumLinCmac As Integer

Dim lsmensaje As String * 39
Dim lsCliAux1 As String
Dim lsCliAux2 As String
Dim psTexto As String

Dim psTextoTC1 As String
Dim psTextoTC2 As String
Dim psMontoSoles As String
Dim sPonderadoSoles As String

'Moneda y Cuenta porque en algunos casos no se manda la moneda... JHVP
If Len(sCodCta) = 18 Then
    nMonedaITF = Mid(sCodCta, 9, 1)
End If

If Len(gsNomAge) > 15 Then
    sNomAge = Mid(sNomAge, 1, 15)
End If

If nMonedaITF = 1 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo"
        lnSaldoDisp = nSaldoDisp
        lnSaldoCnt = nSaldoCont
        
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta"
        lnSaldoDisp = nSaldoDisp - nMonITF
        lnSaldoCnt = nSaldoCont - nMonITF
        
        
    End If
ElseIf nMonedaITF = 2 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo Dolares"
        lnSaldoDisp = nSaldoDisp
        lnSaldoCnt = nSaldoCont
        
        
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta Dolares"
        lnSaldoDisp = nSaldoDisp - nMonITF
        lnSaldoCnt = nSaldoCont - nMonITF
        
    End If
End If

lsNroExt = Str(nNumTran)

ETIQ:

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = gPrnNegritaON
lsNegritaOff = gPrnNegritaOFF
  
nFicSal = FreeFile
Open sLpt For Output As nFicSal

Print #nFicSal, gPrnInicializa;

sSep = 15
sIni = 1
sMax = 33
saux = 5

sFecha = Format$(pdFecha, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(nMonITF, "#,##0.00")
 
lsNomAge = gsNomAge

'Print #nFicSal, gPrnSaltoLinea;
Print #nFicSal, ""
 Print #nFicSal, ""
 Print #nFicSal, ""
If (nVeces Mod 2) = 0 Then
    Print #nFicSal, ""
    Print #nFicSal, ""
End If
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); lsNegritaOn & "CMACICA"; Space(26 + sSep + saux); "CMACICA" & lsNegritaOff

If nMonedaITF = 1 Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If
 
Print #nFicSal, ""
 
Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
lnNumLinCmac = 0

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + saux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(sNomCli); Space(saux + sMax + sSep - Len(sNomCli)); ImpreCarEsp(sNomCli)
Else
    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))

    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)

    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    

    lnCliAux = 1
End If

If Len(Trim(sCodCta)) > 0 Then
    sCodCta = Right("                       " & sCodCta, 18)
    Print #nFicSal, Tab(sIni); "Cuenta:" & sCodCta; Space(8 + sSep + saux); "Cuenta:" & sCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(sTitBol)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
 
Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto
 
If nMonedaITF = 1 Then
    Print #nFicSal, ""  'JVP1
ElseIf nMonedaITF = 2 Then
    
    
    psTextoTC1 = "** Tipo de Cambio SBS"
    If pnTipoPago = 1 Then
        psTextoTC2 = "** ITF Efectivo Soles"
    ElseIf pnTipoPago = 2 Then
        psTextoTC2 = "** ITF Cargo a Cuenta Soles"
    End If
    
    Dim oTC As nTipoCambio
    Set oTC = New nTipoCambio
    Dim pnTCPonderadoSBS As Currency
    
    pnTCPonderadoSBS = oTC.EmiteTipoCambio(pdFecha, TCPonderado)
    If pnTCPonderadoSBS = 0 Then
        psMontoSoles = "0.00"
        sPonderadoSoles = "0.00"
    Else
        psMontoSoles = Format(nMonITF * pnTCPonderadoSBS, "0.00")
        sPonderadoSoles = Format(pnTCPonderadoSBS, "0.0000")
    End If
     
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles
    
    
     
    lnCliAux = lnCliAux + 1
End If
 

Print #nFicSal, lsNegritaOff; 'Desactiva Negrita 'JVP2 JVP3

lsmensaje = "Operacion Origen:"
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff
lsmensaje = psDesOpeOrigen
Print #nFicSal, Tab(sIni); ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje) 'JVP4

Print #nFicSal, "" 'JVP5 JVP6
Print #nFicSal, sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa ''JVP6 JVP7

lnTope = 7 - lnCliAux 'JVP 3-0=3 3-1=2

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)

Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsmensaje = clsGen.GetMensajeBoletas(sCodCta)
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1

For saux = 1 To (lnTope - lnNumLinCmac) 'JVP DE 1 A 3-1=2 DE 1 A 2-1=1
    Print #nFicSal, ""
Next saux

Close nFicSal
Exit Sub
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If

End Sub


Public Sub fgITFImprimeBoleta(ByVal sNomCli As String, ByVal nMonITF As Double, ByVal psDesOpeOrigen As String, _
        ByVal pdFecha As Date, ByVal sLpt As String, Optional nNumTran As Long = 1, _
        Optional sTitBol As String = "IMP. TRANSAC. FINANCIERAS", Optional pnTipoPago As Integer = 1, _
        Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", Optional sGlosa As String, _
        Optional pCancelaImp As Boolean = True, Optional sNomAge As String = "", Optional sCodUser As String = "" _
        , Optional bSaldos As Boolean = False, Optional nSaldoDisp As Double, Optional nSaldoCont As Double, Optional psOpeOrigen As String = "", Optional psAgeDestino As String = "")

If pCancelaImp Then
    Exit Sub 'cancela impresion de itf
End If

Dim psTit As String
Dim sMonto As String
Dim psDesOpe As String
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lnTope As Integer
Dim lsNroExt As String
Dim lnCliAux As Integer
Dim lsNomAge As String
Dim lnSaldoDisp As Double
Dim lnSaldoCnt As Double


Dim lnNumLinCmac As Integer

Dim lsmensaje As String * 39
Dim lsCliAux1 As String
Dim lsCliAux2 As String
Dim psTexto As String

Dim psTextoTC1 As String
Dim psTextoTC2 As String
Dim psMontoSoles As String
Dim sPonderadoSoles As String

'Moneda y Cuenta porque en algunos casos no se manda la moneda... JHVP
If Len(sCodCta) = 18 Then
    nMonedaITF = Mid(sCodCta, 9, 1)
End If

If Len(gsNomAge) > 15 Then
    sNomAge = Mid(sNomAge, 1, 15)
End If

If nMonedaITF = 1 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo"
        lnSaldoDisp = nSaldoDisp
        lnSaldoCnt = nSaldoCont
        
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta"
        lnSaldoDisp = nSaldoDisp - nMonITF
        lnSaldoCnt = nSaldoCont - nMonITF
        
        
    End If
ElseIf nMonedaITF = 2 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo Dolares"
        lnSaldoDisp = nSaldoDisp
        lnSaldoCnt = nSaldoCont
        
        
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta Dolares"
        lnSaldoDisp = nSaldoDisp - nMonITF
        lnSaldoCnt = nSaldoCont - nMonITF
        
    End If
End If

lsNroExt = Str(nNumTran)

ETIQ:

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = gPrnNegritaON
lsNegritaOff = gPrnNegritaOFF
  
nFicSal = FreeFile
Open sLpt For Output As nFicSal

Print #nFicSal, gPrnInicializa;

sSep = 15
sIni = 0
sMax = 33
saux = 5

sFecha = Format$(pdFecha, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(nMonITF, "#,##0.00")
 
lsNomAge = gsNomAge

'Print #nFicSal, gPrnSaltoLinea;


Print #nFicSal, ""
Print #nFicSal, lsNegritaOn; 'Activa Negrita


Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita


Print #nFicSal, Tab(sIni); lsNegritaOn & "CMACICA"; Space(26 + sSep + saux); "CMACICA" & lsNegritaOff

If nMonedaITF = 1 Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If
 
Print #nFicSal, ""
 
Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
lnNumLinCmac = 0

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + saux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(sNomCli); Space(saux + sMax + sSep - Len(sNomCli)); ImpreCarEsp(sNomCli)
Else
    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))

    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)

    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    

    lnCliAux = 1
End If

If Len(Trim(sCodCta)) > 0 Then
    sCodCta = Right("                       " & sCodCta, 18)
    Print #nFicSal, Tab(sIni); "Cuenta:" & sCodCta; Space(8 + sSep + saux); "Cuenta:" & sCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(sTitBol)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
 
Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto
 
If nMonedaITF = 1 Then
    Print #nFicSal, ""  'JVP1
ElseIf nMonedaITF = 2 Then
    
    
    psTextoTC1 = "** Tipo de Cambio SBS"
    If pnTipoPago = 1 Then
        psTextoTC2 = "** ITF Efectivo Soles"
    ElseIf pnTipoPago = 2 Then
        psTextoTC2 = "** ITF Cargo a Cuenta Soles"
    End If
    
    Dim oTC As nTipoCambio
    Set oTC = New nTipoCambio
    Dim pnTCPonderadoSBS As Currency
    
    pnTCPonderadoSBS = oTC.EmiteTipoCambio(pdFecha, TCPonderado)
    If pnTCPonderadoSBS = 0 Then
        psMontoSoles = "0.00"
        sPonderadoSoles = "0.00"
    Else
        psMontoSoles = Format(nMonITF * pnTCPonderadoSBS, "0.00")
        sPonderadoSoles = Format(pnTCPonderadoSBS, "0.0000")
    End If
     
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles
    
    
     
    lnCliAux = lnCliAux + 1
End If

'****PIDIERON MODIFICACION DE SALDO QUE SE PRESENTE DESCONTADO DEL ITF EN LA PRIMERA BOLETA 22/11/2004 mpbr
'If bSaldos = True Then
'         Dim psSaldoDisp As String, psSaldoCont As String
'
'         psSaldoDisp = Format(lnSaldoDisp, "#0.00")
'         psSaldoCont = Format(lnSaldoCnt, "#0.00")
'
'         Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid("Saldo Diponible ", 1, 28))); Space(sMax + 6 - Len(Trim(Mid("Saldo Diponible ", 1, 28))) - Len(psSaldoDisp)); psSaldoDisp; Space(-1 + sSep); ImpreCarEsp(Trim(Mid("Saldo Disponible ", 1, 28))); Space(sMax + 6 - Len(Trim(Mid("Saldo Disponible ", 1, 28))) - Len(psSaldoDisp)); psSaldoDisp
'         Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid("Saldo Contable ", 1, 28))); Space(sMax + 6 - Len(Trim(Mid("Saldo Contable ", 1, 28))) - Len(psSaldoCont)); psSaldoCont; Space(-1 + sSep); ImpreCarEsp(Trim(Mid("Saldo Contable ", 1, 28))); Space(sMax + 6 - Len(Trim(Mid("Saldo Contable ", 1, 28))) - Len(psSaldoCont)); psSaldoCont
'
'End If
     

Print #nFicSal, lsNegritaOff; 'Desactiva Negrita 'JVP2 JVP3

lsmensaje = "Operacion Origen:"
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff
lsmensaje = psDesOpeOrigen
Print #nFicSal, Tab(sIni); ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje) 'JVP4

Print #nFicSal, "" 'JVP5 JVP6
Print #nFicSal, sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa ''JVP6 JVP7

If psAgeDestino <> "" Then
Print #nFicSal, "Destino: " & psAgeDestino & Space(sMax + 5 - Len("Destino: " & psAgeDestino) + sSep) & "Destino: " & psAgeDestino ''JVP6 JVP7
End If

lnTope = 7 - lnCliAux 'JVP 3-0=3 3-1=2

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + saux); ImpreCarEsp(sCodUser)

Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsmensaje = clsGen.GetMensajeBoletas(sCodCta)
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsmensaje); Space(-1 + sSep); ImpreCarEsp(lsmensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1

If psOpeOrigen = "2601" Or psOpeOrigen = "2602" Or psOpeOrigen = "2603" Then
    Print #nFicSal, ""
Else
    For saux = 1 To (lnTope - lnNumLinCmac - 3) 'JVP DE 1 A 3-1=2 DE 1 A 2-1=1
        Print #nFicSal, ""
    Next saux
End If




Close nFicSal
Exit Sub
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If

End Sub

Public Function fgITFGetTitular(psCtaCod As String) As String
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    Sql = " Select Top 1 dbo.PstaNombre(cPersNombre,1) Nombre From ProductoPersona PP" _
        & " Inner Join Persona PE ON PP.cPersCod = PE.cPersCod" _
        & " Where PP.cCtaCod = '" & psCtaCod & "' Order By cPersNombre"
    Set rs = oCon.CargaRecordSet(Sql)
    
    If rs.EOF And rs.BOF Then
        fgITFGetTitular = ""
    Else
        fgITFGetTitular = rs!Nombre
    End If
    
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function fgITFImprimeBoletaCadena(ByVal sNomCli As String, ByVal nMonITF As Double, psDesOpeOrigen As String, pdFecha As Date, sLpt As String, Optional nNumTran As Long = 1, Optional sTitBol As String = "IMP. TRANSAC. FINANCIERAS", Optional pnTipoPago As Integer = 1, Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", Optional sGlosa As String, Optional pCancelaImp As Boolean = True, Optional ByVal sNomAge As String = "", Optional ByVal sCodUser As String = "") As String

If pCancelaImp Then
    Exit Function 'cancela impresion de itf
End If

Dim psTit As String
Dim sMonto As String
Dim psDesOpe As String
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lnTope As Integer
Dim lsNroExt As String
Dim lnCliAux As Integer
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsmensaje As String * 39
Dim lsCliAux1 As String
Dim lsCliAux2 As String
Dim psTexto As String

Dim psTextoTC1 As String
Dim psTextoTC2 As String
Dim psMontoSoles As String
Dim sPonderadoSoles As String

Dim lsBoleta As String

'Moneda y Cuenta porque en algunos casos no se manda la moneda... JHVP
If Len(sCodCta) = 18 Then
    nMonedaITF = Mid(sCodCta, 9, 1)
End If

If Len(gsNomAge) > 15 Then
    gsNomAge = Mid(gsNomAge, 1, 15)
End If

If nMonedaITF = 1 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo"
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta"
    End If
ElseIf nMonedaITF = 2 Then
    If pnTipoPago = 1 Then
        psTexto = "ITF Efectivo Dolares"
    ElseIf pnTipoPago = 2 Then
        psTexto = "ITF Cargo a Cuenta Dolares"
    End If
End If

lsNroExt = Str(nNumTran)

ETIQ:

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = gPrnNegritaON
lsNegritaOff = gPrnNegritaOFF
  
lsBoleta = ""
'lsBoleta = lsBoleta & ""
'lsBoleta = lsBoleta & gPrnInicializa & Chr(10)

lsBoleta = lsBoleta & Chr$(27) & Chr$(64)

lsBoleta = lsBoleta & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
lsBoleta = lsBoleta & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
lsBoleta = lsBoleta & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
lsBoleta = lsBoleta & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
lsBoleta = lsBoleta & Chr$(27) + Chr$(18) ' cancela condensada
lsBoleta = lsBoleta & Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5


sFecha = Format$(pdFecha, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(nMonITF, "#,##0.00")
 
lsNomAge = gsNomAge

'lsBoleta = lsBoleta &  gPrnSaltoLinea;
'lsBoleta = lsBoleta & "." & Chr(10)
'lsBoleta = lsBoleta & "." & Chr(10)
lsBoleta = lsBoleta & lsNegritaOn & Chr(10)  'Activa Negrita
lsBoleta = lsBoleta & Space(sIni) & lsNegritaOn & "CMACICA" & Space(28 + sSep + saux) & "CMACICA" & lsNegritaOff & Chr(10)

If nMonedaITF = 1 Then
    lsBoleta = lsBoleta & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr(10)
Else
    lsBoleta = lsBoleta & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr(10)
End If
lsBoleta = lsBoleta & Chr$(27) & Chr$(67) & Chr$(22)
lsBoleta = lsBoleta & "" & Chr(10)
 
lsBoleta = lsBoleta & lsNegritaOff & Chr(10)  'Desactiva Negrita
lnNumLinCmac = 0

lsBoleta = lsBoleta & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Chr(10)

lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)

If lnCliAux = 0 Then
    If saux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + saux)
    lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(sNomCli) & Space(saux + sMax + sSep - Len(sNomCli)) & ImpreCarEsp(sNomCli) & Chr(10)
Else
    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))

    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)

    lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & Chr(10)
    lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & Chr(10)

    lnCliAux = 1
End If

If Len(Trim(sCodCta)) > 0 Then
    sCodCta = Right("                       " & sCodCta, 18)
    lsBoleta = lsBoleta & Space(sIni) & "Cuenta:" & sCodCta & Space(14 + sSep + saux) & "Cuenta:" & sCodCta & Chr(10)
Else
    lsBoleta = lsBoleta & "" & Chr(10)
End If

psTit = Trim(sTitBol)
psTit = CentrarCadena(psTit, 28)
lsBoleta = lsBoleta & lsNegritaOn & Chr(10)  'Activa Negrita
lsBoleta = lsBoleta & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr(10)
 
lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)) & sMonto & Chr(10)
 
If nMonedaITF = 1 Then
    lsBoleta = lsBoleta & "" & Chr(10)   'JVP1
ElseIf nMonedaITF = 2 Then
    psTextoTC1 = "** Tipo de Cambio SBS"
    If pnTipoPago = 1 Then
        psTextoTC2 = "** ITF Efectivo Soles"
    ElseIf pnTipoPago = 2 Then
        psTextoTC2 = "** ITF Cargo a Cuenta Soles"
    End If
    
    Dim oTC As nTipoCambio
    Set oTC = New nTipoCambio
    Dim pnTCPonderadoSBS As Currency
    
    pnTCPonderadoSBS = oTC.EmiteTipoCambio(pdFecha, TCPonderado)
    If pnTCPonderadoSBS = 0 Then
        psMontoSoles = "0.00"
        sPonderadoSoles = "0.00"
    Else
        psMontoSoles = Format(nMonITF * pnTCPonderadoSBS, "0.00")
        sPonderadoSoles = Format(pnTCPonderadoSBS, "0.0000")
    End If
     
    lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)) & sPonderadoSoles & Space(-1 + sSep) & ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)) & sPonderadoSoles & Chr(10)
    lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)) & psMontoSoles & Space(-1 + sSep) & ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)) & psMontoSoles & Chr(10)
     
    lnCliAux = lnCliAux + 1
End If

lsBoleta = lsBoleta & lsNegritaOff & Chr(10)  'Desactiva Negrita 'JVP2 JVP3

lsmensaje = "Operacion Origen:"
lsBoleta = lsBoleta & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & ImpreCarEsp(lsmensaje) & lsNegritaOff & Chr(10)
lsmensaje = psDesOpeOrigen
lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & ImpreCarEsp(lsmensaje) & Chr(10) 'JVP4

lsBoleta = lsBoleta & "" & Chr(10)  'JVP5 JVP6
 
lsBoleta = lsBoleta & sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa & Chr(10)  ''JVP6 JVP7
     
lnTope = 7 - lnCliAux 'JVP 3-0=3 3-1=2

lsBoleta = lsBoleta & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr(10)
lsBoleta = lsBoleta & Space(sIni) & ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & ImpreCarEsp(sCodUser) & Chr(10)

Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsmensaje = clsGen.GetMensajeBoletas(sCodCta)
lsBoleta = lsBoleta & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & ImpreCarEsp(lsmensaje) & lsNegritaOff & Chr(10)

lnNumLinCmac = lnNumLinCmac + 1

'For saux = 1 To (lnTope - lnNumLinCmac) 'JVP DE 1 A 3-1=2 DE 1 A 2-1=1
For saux = 1 To 4
    lsBoleta = lsBoleta & "" & Chr(10)
Next saux

fgITFImprimeBoletaCadena = lsBoleta

Exit Function
Error:
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function

