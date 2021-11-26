VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContImprimir"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EE84101F4"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim vsAgeNom As String
Dim vsEmpNom As String
Dim vsFecSis As String
Dim lsEmprLogo As String

Public Event BarraShow(pnMax)
Public Event BarraProgress(value, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
Public Event BarraClose()



Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing
    
    'Dim oCons As New COMDConstSistema.NCOMConstSistema
    lsEmprLogo = "CMAC MAYNAS S.A." 'oCons.LeeConstSistema(gConstSistNombreAbrevCMAC)
    gsNomCmac = "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS" 'oCons.LeeConstSistema(gConstSistCMACNombreCompleto)
    'Set oCons = Nothing
    
      
'    sLpt = "LPT1"
'    BON = PrnSet("B+")
'    BOFF = PrnSet("B-")
'    CON = PrnSet("C+")
'    COFF = PrnSet("C-")
    
    'Set oIni = Nothing
End Sub

Private Function CabeceraPlanillaRendicion(ByVal pnTipoArendir As ArendirTipo, ByVal psTitulo As String, ByRef nLin As Integer, ByRef P As Integer, ByVal pnLinPage As Integer, _
                                           ByVal pnColPage As Integer, ByVal psNomCmac As String, _
                                           ByVal pdFecsis As Date, ByVal psNroDocArendir As String, ByVal psDescDocAtend As String, _
                                           ByVal psNroDocAtend As String, ByVal psDocAtencFecha As String, _
                                           ByVal psSimbolo As String, ByVal ImporteARendir As Currency, _
                                           ByVal psAgeCod As String, ByVal psAgeDesc As String, _
                                           ByVal psAreaCod As String, ByVal psAreaDesc As String, _
                                           ByVal psPersCod As String, ByVal psPersNombre As String, _
                                           ByVal psAreaCodARendir As String, ByVal psAgeCodArendir As String, _
                                           ByVal psAreaDescArendir As String, ByVal psAgeDescArendir As String, _
                                           Optional pbCabRend As Boolean = True) As String
Dim lsTexto As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String

If nLin > pnLinPage - 4 Then
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
   If P > 0 Then lsTexto = lsTexto & oImpresora.gPrnSaltoPagina
   P = P + 1
   lsTexto = lsTexto + ImpreFormat(psNomCmac, 60) & pdFecsis & " - " & Format(Time, "hh:mm:ss") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + " CONTABILIDAD    " & Space(55) & "Pag. " & Format(P, "000") + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + BON & Centra(psTitulo, pnColPage) & BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
'   If pnTipoArendir = gArendirTipoCajaGeneral Then
        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCod & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & Mid(psDescDocAtend & Space(18), 1, 18) & ": " & BON & psNroDocAtend & BOFF & "       Fecha : " & BON & psDocAtencFecha & BOFF & _
            Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
'   Else
'        lsTexto = lsTexto + CON & "  Recibo de A rendir: " & BON & psNroDocArendir & BOFF + Space(15) & "Monto : " & BON & Right(Space(18) & psSimbolo & " " & Format(ImporteARendir, "#,#0.00"), 18) & BOFF + oImpresora.gPrnSaltoLinea
'        lsTexto = lsTexto + "  A Rendir de.......: " & BON & psAreaCodARendir + psAgeCodArendir & " " & psAreaDescArendir + " " + psAgeDescArendir & BOFF + oImpresora.gPrnSaltoLinea
'   End If
   lsTexto = lsTexto + "  Usuario...........: " & BON & psPersCod & " " & PstaNombre(psPersNombre) & BOFF + oImpresora.gPrnSaltoLinea
   lsTexto = lsTexto + "  Area Funcional....: " & BON & psAreaCod & " " & psAreaDesc & BOFF + oImpresora.gPrnSaltoLinea
   If psAgeCod <> "" Then
        lsTexto = lsTexto + "  Agencia       ....: " & BON & psAgeCod & " " & psAgeDesc & BOFF + oImpresora.gPrnSaltoLinea
   End If
   lsTexto = lsTexto + oImpresora.gPrnSaltoLinea
   If pbCabRend Then
        lsTexto = lsTexto + " ====================================================================================================================================" & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "        DOCUMENTOS SUSTENTATORIOS          " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  ITEM  --------------------------------  PROVEEDOR                                 DETALLE                                  IMPORTE " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "        Fecha      Tpo  Número                                                                                                       " & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + " ------------------------------------------------------------------------------------------------------------------------------------" & COFF + oImpresora.gPrnSaltoLinea
    Else
        lsTexto = lsTexto + COFF + BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & oImpresora.gPrnSaltoLinea
        lsTexto = lsTexto + String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End If
    nLin = 15
End If
CabeceraPlanillaRendicion = lsTexto
End Function

Public Function ImprimeNotaAbono(pdFecha As String, pnImporte As Currency, psGlosa As String, Optional lsCodCta As String, Optional psPersona As String = "", Optional sDocPagado As String = "") As String
Dim sTexto As String, sImpre As String
Dim nLen As Integer, N As Integer
Dim nLin As Integer
Dim sTextoAux As String

    sImpre = oImpresora.gPrnInicializa + oImpresora.gPrnTamLetra10CPI
    Linea sImpre, PrnSet("Esp", 6)
    Linea sImpre, PrnSet("B+") + Space(48), 7
    Linea sImpre, PrnSet("Esp", 1)
    Linea sImpre, Space(12) + IIf(Len(psPersona) > 45, CON, "") + psPersona + COFF, 2
    Linea sImpre, Space(52) & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Mid(pdFecha, 7, 4)
    Linea sImpre, PrnSet("EspN")
    Linea sImpre, PrnSet("Esp", 4)
    Linea sImpre, Space(48), 5
    Linea sImpre, Space(50) & lsCodCta
    Linea sImpre, PrnSet("EspN"), 3
    If Len(lsCodCta) = 12 Then
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional, gcMN, IIf(Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13)
    Else
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional, gcMN, IIf(Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera, gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13)
    End If
    sTextoAux = Trim(psGlosa) & " " & Trim(sDocPagado)
    nLen = 6
    'sTexto = Space(8) & JustificaTexto(gcGlosa & oImpresora.gPrnSaltoLinea & sDocPagado, 44 - nLen)
    sTexto = IIf(Len(sTextoAux) > 100, Space(4), Space(7)) & JustificaTexto(IIf(Len(sTextoAux) > 100, PrnSet("C+"), "") + sTextoAux, IIf(Len(sTextoAux) > 100, 61, 44) - nLen, IIf(Len(sTextoAux) > 100, 4, 0)) & PrnSet("C-")
    N = 0
    Do While True
       nLin = nLin + 1
       N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
       If N > 0 Then
          sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
          sTexto = Mid(sTexto, N + 1, Len(sTexto))
       End If
       If N = 0 Then
          sImpre = sImpre & sTexto & oImpresora.gPrnSaltoLinea
          Exit Do
       End If
    Loop
    Linea sImpre, "", 8 - nLin
    If Len(lsCodCta) = 12 Then
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 6, 1) = "1", gcMN, IIf(Mid(lsCodCta, 6, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13)
    Else
        Linea sImpre, Space(48) & IIf(Mid(lsCodCta, 9, 1) = "1", gcMN, IIf(Mid(lsCodCta, 9, 1) = "2", gcME, "")) & Right(Space(13) & Format(pnImporte, gsFormatoNumeroView), 13)
    End If
    ImprimeNotaAbono = ImpreCarEsp(sImpre) + PrnSet("B-")
End Function

Public Function ImprimeNotaCargoAbono(ByVal psNroNota As String, ByVal psGlosa As String, _
                                        ByVal pnImporte As Currency, _
                                        ByVal psPersNombre As String, ByVal psPersDireccion As String, _
                                        ByVal psPersUbiGeo As String, ByVal pdFecha As Date, ByVal psMoneda As Moneda, _
                                        ByVal psNroCtaAhorros As String, ByVal pnTpoDoc As TpoDoc, ByVal psAgeDesc As String, _
                                        ByVal psCodUser As String)
Dim lsTexto As String
Dim lsTitulo As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String
Dim lnColIzq As Long
Dim lnCol As Long
Dim lsGlosa As String
Dim lsGlosaLin As String
Dim Char As String
Dim lnPos As Integer
Dim lnLineas As Integer
Dim lnImporteCargo As Currency
Dim lnImporteAbono As Currency
Dim lsTextoGlosa As String
Dim lsTextoTotal As String
Dim I As Integer
Dim lnLineasGlosa As Integer
lnColIzq = 13
lnCol = 78

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

If pnTpoDoc = TpoDocNotaCargo Then
    lsTitulo = "NOTA DE CARGO N° " & psNroNota
    lnImporteCargo = pnImporte
    lnImporteAbono = 0
    lsTextoTotal = " TOTAL CARGADO "
Else
    lsTitulo = "NOTA DE ABONO N° " & psNroNota
    lnImporteAbono = pnImporte
    lnImporteCargo = 0
    lsTextoTotal = " TOTAL ABONADO "
End If
lsTexto = CON + BON + PrnSet("I+") + ImpreFormat(lsTitulo, lnCol, 50) & ImpreFormat(lsTitulo, lnCol, 10) + PrnSet("I-") + BOFF + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(Trim(psAgeDesc), lnCol, 3) & ImpreFormat(Trim(psAgeDesc), lnCol, lnColIzq) + oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
lsTexto = lsTexto + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, 3) + ImpreFormat(CentrarCadena(String(33, "=") + " Concepto " + String(35, "="), lnCol), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lnLineas = lnLineas + 1
'Impresion de la glosa
lsGlosa = JustificaTexto(psGlosa, lnCol - 6)
lsGlosaLin = "": lnPos = 0: Char = ""
lnLineasGlosa = 0
For I = 1 To Len(lsGlosa)
    Char = Mid(lsGlosa, I, 1)
    lsGlosaLin = lsGlosaLin & Char
    If Char = oImpresora.gPrnSaltoLinea Or I = Len(lsGlosa) Then
        lsGlosaLin = Mid(lsGlosaLin, 1, IIf(I = Len(lsGlosa), Len(lsGlosaLin), Len(lsGlosaLin) - 1))
        lsTextoGlosa = lsTextoGlosa & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq - 7) & ImpreFormat(Trim(lsGlosaLin), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
        lnLineas = lnLineas + 1
        lnLineasGlosa = lnLineasGlosa + 1
        If lnLineasGlosa = 3 Then
            Exit For
        End If
        lsGlosaLin = ""
    End If
Next
lsTexto = lsTexto + lsTextoGlosa
lsTexto = lsTexto + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, 3) + ImpreFormat(String(50, "_") & " CARGOS " & String(10, "_") & " ABONOS " & String(14, "_"), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
If lnImporteCargo > 0 Then
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(lnImporteCargo, 14, 2, True) & ImpreFormat(0, 15, 2, True), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
Else
    lsTexto = lsTexto + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, 3) + ImpreFormat(ImpreFormat("IMPORTE :", 40) & ImpreFormat(0, 14, 2, True) & ImpreFormat(lnImporteAbono, 15, 2, True), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
End If
lsTexto = lsTexto + ImpreFormat(String(lnCol, "_"), lnCol, 3) + ImpreFormat(String(lnCol, "_"), lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat("Sr(es) :", lnCol, 3) + ImpreFormat("Sr(es) :", lnCol, lnColIzq) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + BON + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) + ImpreFormat(PstaNombre(psPersNombre), lnCol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) + ImpreFormat(psPersDireccion, lnCol, lnColIzq + 5) & oImpresora.gPrnSaltoLinea
lsTexto = lsTexto + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + ImpreFormat(psPersUbiGeo, lnCol, lnColIzq + 5) + BOFF + oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(15, "-"), lnCol + 4, 3) _
            + ImpreFormat(String(3, "-") & "FECHA" & String(13, "-") & " MONEDA " & String(5, "-") & " CUENTA " & String(15, "-") + BON + lsTextoTotal + BOFF & String(8, "-"), lnCol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, 3) _
            + ImpreFormat(String(3, " ") & Format(pdFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & String(3, " ") & IIf(psMoneda = gMonedaNacional, "SOLES", "DOLARES") & String(3, " ") & psNroCtaAhorros + BON + ImpreFormat(pnImporte, 15, 2, True) + BOFF + ImpreFormat(psCodUser, 5, 4), lnCol + 4, lnColIzq) & oImpresora.gPrnSaltoLinea

lsTexto = lsTexto + ImpreFormat(String(lnCol, "-"), lnCol, 3) + ImpreFormat(String(lnCol, "-"), lnCol, lnColIzq) + COFF + oImpresora.gPrnSaltoLinea
ImprimeNotaCargoAbono = lsTexto
End Function

Public Sub Inicio(psEmpNom As String, psAgeNom As String, psFecSis As String)
vsAgeNom = psAgeNom
vsEmpNom = psEmpNom
vsFecSis = psFecSis
End Sub

Private Function CabeceraRepo(ByVal psSimbolo As String, ByVal pnLinPage As Integer, ByVal psNomCmact As String, ByVal pdFecsis As Date, _
                              ByVal pnColPage As Integer, ByVal pbValido As Boolean, _
                              ByVal psNroCH As String, ByVal pnNroProc As Integer, ByVal psCHDesc As String, _
                              ByVal pdFechaApertReembolso As Date, ByVal pnMontoAsignado As Currency, _
                              ByVal psPersCod As String, ByVal psPersNombre As String, _
                              ByRef nLin As Integer, ByRef P As Integer) As String
Dim sCabe As String
sCabe = ""
If nLin > pnLinPage - 6 Then
   If P > 0 Then sCabe = sCabe & oImpresora.gPrnSaltoPagina
   P = P + 1
   Linea sCabe, CON + ImpreFormat(psNomCmact, 60, 0) & Space(42) & pdFecsis & " - " & Format(Time, "hh:mm:ss")
   Linea sCabe, ImpreFormat("CONTABILIDAD", 60, 0) & Space(48) & "Pag. " & Format(P, "000") & COFF
   Linea sCabe, BON & Centra(" DOCUMENTOS SUSTENTATORIOS DE CAJA CHICA ", pnColPage) & BOFF
   Linea sCabe, CON
   If pbValido = False Then
      Linea sCabe, BON & Centra("*** REPORTE NO VALIDO PARA RENDICION EN CONTABILIDAD ***", 80) & BOFF, 2
   End If
   If P = 1 Then
      Linea sCabe, " Area Responsable    : " & BON & ImpreFormat("[" & psNroCH & "] " & psCHDesc, 40, 0) & ImpreFormat("Caja Chica N° [" & pnNroProc & "]", 28) & BOFF & _
                    " Apertura/Reembolso : " & BON & pdFechaApertReembolso & BOFF
      Linea sCabe, " Persona Responsable : " & BON & Left("[" & psPersCod & "] " & psPersNombre & Space(70), 70) & BOFF & _
               "     Monto Asignado : " & psSimbolo & " " & BON & Format(pnMontoAsignado, "#,#0.00") & BOFF
      nLin = 10
   Else
      nLin = 8
   End If
   Linea sCabe, " ===================================================================================================================================="
   Linea sCabe, "   Item  Fecha de   Comprobante          Proveedor                                 Detalle                                    Monto  "
   Linea sCabe, "         Emisión    Tpo  Número                                                                         "
   Linea sCabe, " ------------------------------------------------------------------------------------------------------------------------------------" & COFF
End If
CabeceraRepo = sCabe
End Function

Public Function ImpreAsiento(ByVal pnColPage As Integer, ByVal psSimbolo As String, _
                            ByVal psAsiento As String, ByVal pnTotalDebe As Currency, _
                            ByVal pnTotalHaber As Currency, lPiePage As Boolean, Optional lsArea As String = "CAJA") As String
Dim sImpre As String, nLon  As String
sImpre = sImpre & oImpresora.gPrnSaltoLinea
' ASIENTO CONTABLE
sImpre = sImpre & BON & "ASIENTO CONTABLE EN " & IIf(psSimbolo = "S/.", "M.N. :", "M.E. :") & BOFF & oImpresora.gPrnSaltoLinea
sImpre = sImpre & String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & "  " & BON & "Cuenta Contable" & Space(43) & "DEBE       HABER" & BOFF & oImpresora.gPrnSaltoLinea
sImpre = sImpre & String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & CON & psAsiento & COFF     'Detalle de Asiento
sImpre = sImpre & String(pnColPage - 1, "-") & oImpresora.gPrnSaltoLinea
sImpre = sImpre & CON
sImpre = sImpre & Space(75) & "TOTALES        " & psSimbolo & Right(Space(15) & Format(pnTotalDebe, "###,###,##0.00"), 15) & "   " & psSimbolo & Right(Space(15) & Format(pnTotalHaber, "###,###,##0.00"), 15) & oImpresora.gPrnSaltoLinea
sImpre = sImpre & COFF
sImpre = sImpre & String(pnColPage - 1, "=") & oImpresora.gPrnSaltoLinea

If lPiePage Then
   sImpre = sImpre & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
   sImpre = sImpre & BON
   sImpre = sImpre & Space(5) & "_____________________                           ______________________" & oImpresora.gPrnSaltoLinea
   lsArea = "Vo Bo " & lsArea
   nLon = Int((21 - Len(lsArea)) / 2)
   sImpre = sImpre & Space(5) & Space(nLon) & lsArea & Space(nLon) & "                              Vo Bo CONTABILIDAD" & BOFF & oImpresora.gPrnSaltoLinea
End If
ImpreAsiento = sImpre
End Function

Private Function ImprimeOperacionesCab(nPag As Integer) As String
Dim lsTexto As String
Linea lsTexto, Justifica("CMACT", 15) + CentrarCadena("Pagina Nro " + Format(nPag, "00"), 185) + CStr(Date)
Linea lsTexto, "Sede Principal " + Space(185) + CStr(Time)
Linea lsTexto, CentrarCadena("Listado de Operaciones", 215)
Linea lsTexto, String(215, "=")
Linea lsTexto, "            O P E R A C I O N                                     C U E N T A   C O N T A B L E                                       O B J E T O S                                             D O C U M E N T O S"
Linea lsTexto, "---------------------------------------------   ---------------------------------------------------------------------    ---------------------------------------------------------------   ----------------------------"
Linea lsTexto, "CODIGO DESCRIPCION                              CUENTA      DESCRIPCION                            CLASE  TIPO           CODIGO               DESCRIPCION                                  TIPO    DENOMINACION  "
Linea lsTexto, String(215, "-")
ImprimeOperacionesCab = lsTexto
End Function

Public Function ImprimePlanContable(prs As Recordset, psNomTabla As String, pnLinPage As Integer) As String
' Variables de Sangrado de Cta
Dim nLongC, nLongT As Integer, nLongD As Integer, nNivel As Integer
Dim aSpace(1 To 20, 1 To 2) As Integer, nPos As Integer, Espacios As Integer
' Variables de Impresión
Dim lnCarLin, lnPagNro As Integer
Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
Dim lnPiePag, lnLinea As Integer
Dim sTexto As String

For nPos = 1 To 20
    aSpace(nPos, 1) = 0
Next
nNivel = 0
nLongC = 0

lnCarLin = 90: lnLinea = 70
lnPiePag = 5: lnPagNro = 0

   If Trim(psNomTabla) = "CtaCont" Then
      lsTitRep = "P L A N   C O N T A B L E   G E N E R A L"
   Else
      lsTitRep = "P L A N   C O N T A B L E   B A S E   S B S"
   End If
  lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
sTexto = ""
Do While Not prs.EOF
   DoEvents
   ' Sangra Código
   nLongT = Len(Trim(prs!cCtaContCod))
   Espacios = nLongT
   nLongC = Len(Trim(prs!cCtaContCod))
   nLongD = Len(Trim(prs!cCtaContDesc))
   
   'Salto de Pagina si cumple condicion
   If lnLinea > pnLinPage - lnPiePag - 1 Then
        If lnPagNro > 0 Then
           Linea sTexto, String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina
        End If
        lnPagNro = lnPagNro + 1
        ' Definición de Cabecera 1
        lsCabe01 = vsEmpNom
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(lnPagNro)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(vsFecSis, gsFormatoFechaView)
        ' Definición de Cabecera 2
        lsCabe02 = "CONTABILIDAD" & "-" & vsAgeNom
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
        
   '    Cabecera de Impresión
        Linea sTexto, lsCabe01
        Linea sTexto, lsCabe02

        Linea sTexto, UCase(lsTitRep), 2
        Linea sTexto, String(lnCarLin, "-")
        Linea sTexto, "   CODIGO   " & String(15, " ") & "DESCRIPCION" & String(12, " ") & "   "
        Linea sTexto, String(lnCarLin, "-"), 2
        lnLinea = 9
    End If
    Linea sTexto, Space(Espacios) & Trim(prs!cCtaContCod) _
        & " " & Left(prs!cCtaContDesc, nLongD + 1)
    lnLinea = lnLinea + 1
    prs.MoveNext
Loop
Linea sTexto, String(lnCarLin, "-")             ' hasta 160 caracteres x línea
ImprimePlanContable = sTexto
End Function

Public Function ImprimeArqueo(ByVal psNomCmact As String, ByVal pdFecha As Date, _
                             ByVal psMovNro As String, ByVal psOpeCod As String, _
                             ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                             ByVal rsTotales As ADODB.Recordset, ByVal rsOP As ADODB.Recordset, _
                             ByVal psAreaAgeCh As String, ByVal psCajaChicaNomb As String, _
                             ByVal psCodPersResp As String, ByVal psResponsable As String, _
                             ByVal psAgeCod As String, ByVal psAgencia As String, _
                             ByVal psCodPers As String, ByVal psPersNombre As String, _
                             ByVal psHoraIni As String, ByVal psHoraFin As String, _
                             ByVal pnNumPag As Integer, ByVal pnNumCol As Integer) As String
            

Dim lsTexto As String
Dim lnTotalBill As Currency
Dim lnTotalMon As Currency
Dim lsCab1 As String
Dim lsCab2 As String
Dim lsTitulo2 As String
Dim lnNumPag As Integer
On Error GoTo ImprimeArqueoErr

lnTotalBill = 0
lnTotalMon = 0
lsTexto = PrnSet("MI", 10)
lsCab1 = CON + ImpreFormat("Resp. Caja   : [" & psCodPersResp & "] " + psResponsable, (pnNumCol / 2) + 65) + BON + ImpreFormat("Inicio  :" & psHoraIni, 25) + COFF + BOFF
lsCab2 = CON + ImpreFormat("Resp. Arqueo : [" & psCodPers & "] " + psPersNombre, (pnNumCol / 2) + 65) + BON + ImpreFormat("Término :" & psHoraFin, 25) + COFF + BOFF
lsTitulo2 = PrnSet("I+") + "Caja Chica N° [" & psAreaAgeCh & "] " & psCajaChicaNomb + PrnSet("I-")
lnNumPag = 0
Linea lsTexto, CabeRepo(psNomCmact, psAgencia, "CONTABILIDAD", IIf(Mid(psOpeCod, 3, 1) = "1", "SOLES", "DOLARES"), Format(pdFecha, gsFormatoFechaView), "A R Q U E O   D E   C A J A   C H I C A", lsTitulo2, "", "", lnNumPag, pnNumCol)
Linea lsTexto, lsCab1
Linea lsTexto, lsCab2
Linea lsTexto, ""
Linea lsTexto, BON + ImpreFormat("FONDOS RECONTADOS ", pnNumCol)
Linea lsTexto, ImpreFormat("I. BILLETES ", pnNumCol)
Linea lsTexto, String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("VALOR", 27) + ImpreFormat("CANTIDAD", 12, 0, True) & ImpreFormat("IMPORTE", 15, 2, True)
Linea lsTexto, String(pnNumCol - 20, "-") + BOFF
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                Linea lsTexto, ImpreFormat(rsBillIngreso(0), 20) + ImpreFormat(rsBillIngreso(1), 10, 0, True) & ImpreFormat(rsBillIngreso(2), 15, 2, True)
                lnTotalBill = lnTotalBill + rsBillIngreso(2)
                rsBillIngreso.MoveNext
            Loop
        End If
        'rsBillIngreso.Close: Set rsBillIngreso = Nothing
    End If
End If
Linea lsTexto, BON + String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("TOTAL :", 20) + ImpreFormat("", 10, 0, True) & ImpreFormat(lnTotalBill, 15, 2, True) + BOFF

Linea lsTexto, ""
Linea lsTexto, BON + ImpreFormat("II. MONEDAS ", pnNumCol)
Linea lsTexto, String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("VALOR", 27) + ImpreFormat("CANTIDAD", 12, 0, True) & ImpreFormat("IMPORTE", 15, 2, True)
Linea lsTexto, String(pnNumCol - 20, "-") + BOFF

If Not rsMonIng Is Nothing Then
    If rsMonIng.State = adStateOpen Then
        If Not rsMonIng.EOF And Not rsMonIng.BOF Then
            Do While Not rsMonIng.EOF
                Linea lsTexto, ImpreFormat(rsMonIng(0), 20) + ImpreFormat(rsMonIng(1), 10, 0, True) & ImpreFormat(rsMonIng(2), 15, 2, True)
                lnTotalMon = lnTotalMon + rsMonIng(2)
                rsMonIng.MoveNext
            Loop
        End If
        'rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
Linea lsTexto, BON + String(pnNumCol - 20, "-")
Linea lsTexto, ImpreFormat("TOTAL :", 20) + ImpreFormat("", 10, 0, True) & ImpreFormat(lnTotalMon, 15, 2, True) + BOFF
Linea lsTexto, ""
If Not rsOP Is Nothing Then
    If rsOP.State = adStateOpen Then
        If Not rsOP.EOF And Not rsOP.BOF Then
        Linea lsTexto, BON + ImpreFormat("III. ORDENES DE PAGO PENDIENTE DE COBRO", pnNumCol)
        Linea lsTexto, String(pnNumCol - 20, "-")
        Linea lsTexto, ImpreFormat("N°", 10) + ImpreFormat("ORDEN", 20) & ImpreFormat("IMPORTE", 15, 2, True)
        Linea lsTexto, String(pnNumCol - 20, "-") + BOFF
            Do While Not rsOP.EOF
                Linea lsTexto, ImpreFormat(rsOP.Bookmark, 10) + ImpreFormat(rsOP!Orden, 20) & ImpreFormat(rsOP!Importe, 15, 2, True)
                rsOP.MoveNext
            Loop
        End If
        rsOP.Close: Set rsOP = Nothing
    End If
End If
If Not rsTotales Is Nothing Then
    If rsTotales.State = adStateOpen Then
        If Not rsTotales.EOF And Not rsTotales.BOF Then
            Linea lsTexto, "", 1
            Do While Not rsTotales.EOF
                If rsTotales.Bookmark = 7 Or rsTotales.Bookmark = 9 Then
                    lsTexto = lsTexto + BON
                End If
                Linea lsTexto, ImpreFormat(rsTotales(0), 40) + ImpreFormat(rsTotales(1), 15, 2, True) + BOFF
                rsTotales.MoveNext
            Loop
        End If
    End If
End If
Linea lsTexto, "", 2
Linea lsTexto, "Siendo las " & psHoraFin & " se dio por terminado el arqueo de Caja,"
Linea lsTexto, "devolviendose al responsable " & psResponsable & " todo el dinero y documentos de Caja,"
Linea lsTexto, "tal como se nos entregó para efecto del arqueo."
Linea lsTexto, "Se firma la presente en señal de conformidad"
'Linea lsTexto, "", 2
Linea lsTexto, ImprePiePag(pnNumCol, "69")

ImprimeArqueo = lsTexto

Exit Function
ImprimeArqueoErr:
    Err.Raise vbObjectError + 100, "ImprimeArqueo", Err.Description & " Error ImprimeArqueo"
End Function

Private Function ImprimeLibroDiarioCab(nLi As Integer, P As Integer, dFecIni As Date, dFecFin As Date, nNro As Currency, sFec As String, sAge As String, sTpo As String, pnLinPage As Integer) As String
Dim sCadCab As String
If pnLinPage - 6 < nLi Then
   Linea sCadCab, COFF & BON & Cabecera("L I B R O   D I A R I O ", P, "", 80, "", Format(dFecFin, gsFormatoFechaView), gsNomCmac), 2
   Linea sCadCab, CON & Space(5) + "CODIGO" + Space(15) + "NOMBRE" + Space(38) + "DEBE" + Space(8) + "HABER", 2
   Linea sCadCab, "Documento: " & Format(nNro, "#####0") & "     Tipo de Asiento: " & sTpo & "      Agencia: " & sAge & "      Fecha: " & sFec & BOFF, 2
   nLi = 9
End If
ImprimeLibroDiarioCab = sCadCab
End Function

Private Function ImprimeLibroMayorCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim K As Integer, nObjs As Integer
Dim sCabe As String
If nLin > pnLinPage - 6 Then
   Linea sCabe, BON & Cabecera(" L I B R O    M A Y O R ", P, "", 80, , Format(dFechaAl, gsFormatoFechaView), lsEmprLogo), 0
   Linea sCabe, BON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80) & BOFF, 2
   Linea sCabe, CON & "FECHA       DOC   ASIENTO            DEBE         HABER        SALDO FINAL   DESCRIPCION" & COFF, 2
   Linea sCabe, "CUENTA: " & Justifica(sCtaCod, 22) & " " & sCtaDes & BOFF, 2
   nLin = 10
End If
ImprimeLibroMayorCab = sCabe
End Function

Public Function ImprimeMayorCta(psCtaCod As String, psCtaDes As String, pdFechaDel As Date, pdFechaAl As Date, pnImporte As Currency, pnSaldoIniME As Currency, pnSaldoFinalME As Currency, psFiltro As String, pnLinPage As Integer, Optional pbVerPersona As Boolean = False) As String
Dim sTexto As String
Dim nRow As Integer
Dim sMov As String
Dim sDoc As String
Dim N As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nHaberD As Currency
Dim nSaldo  As Currency, nSaldoIni As Currency
Dim sFecha  As String
Dim rs      As ADODB.Recordset
Dim rsCta   As ADODB.Recordset
Dim nLin    As Integer
Dim P       As Integer
Dim sTipoCta As String

nLin = pnLinPage
P = 0
Dim oCont As New NContAsientos
Dim oSdo  As New NCtasaldo
Set oSdo = Nothing

Set rsCta = oCont.GetMayorCuenta(psCtaCod, Format(pdFechaDel, "yyyymmdd"), Format(pdFechaAl, "yyyymmdd"), pnImporte, psFiltro, , , pbVerPersona)
Set oCont = Nothing
If rsCta.EOF Then
   nSaldo = oSdo.GetCtaSaldo(psCtaCod + "%", Format(pdFechaDel - 1, gsFormatoFecha))
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   Linea sTexto, oImpresora.gPrnCondensadaON & Space(87) & "SALDO INICIAL : " & PrnVal(nSaldo, 16, 2)
   Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnBoldOFF, 1
   ImprimeMayorCta = sTexto
   Set oCont = Nothing
   Exit Function
End If
Dim oCta  As New DCtaCont

Set rs = oCta.CargaCtaContClase(psCtaCod)
Set oCta = Nothing
If Not rs.EOF Then
   sTipoCta = Trim(rs!cCtaCaracter)
Else
   sTipoCta = "D"
End If

RaiseEvent BarraShow(rsCta.RecordCount)
sTexto = ""
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
nLin = nLin + 2
nDebe = 0: nHaber = 0
N = 1

sFecha = Mid(rsCta!cMovNro, 1, 8)
nSaldo = oSdo.GetCtaSaldo(rsCta!cCtaContCod, Format(pdFechaDel - 1, gsFormatoFecha))
nSaldoIni = nSaldo
Do While Not rsCta.EOF
   DoEvents
   RaiseEvent BarraProgress(rsCta.Bookmark, "Mayor de Cuenta Contable : " & psCtaCod, "", "Generando Reporte... ", vbBlue)
   sMov = rsCta!cMovNro
   If sFecha <> Mid(sMov, 1, 8) Then
      Linea sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnBoldOFF, 2, nLin
      nDebeD = 0: nHaberD = 0
      sFecha = Mid(rsCta!cMovNro, 1, 8)
   End If
   If sTipoCta = "D" Then
      nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
   Else
      nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
   End If
   
   Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
   'Linea Detalle de Reporte
   sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & Justifica(rsCta!dDocFecha, 10) & " "
   sMov = rsCta!cMovNro
   Linea sTexto, oImpresora.gPrnCondensadaON & " " & Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4) & "  " _
          & " " & sDoc & Mid(Replace(Replace(rsCta!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, "") & Space(100), 1, 42) & " " _
          & PrnVal(rsCta!nDebe, 16, 2, False) & "  " & PrnVal(rsCta!nHaber, 16, 2, False) _
          & oImpresora.gPrnCondensadaOFF
   nLin = nLin + 1
   nDebeD = nDebeD + rsCta!nDebe
   nDebe = nDebe + rsCta!nDebe
   nHaberD = nHaberD + rsCta!nHaber
   nHaber = nHaber + rsCta!nHaber
   rsCta.MoveNext
   If rsCta.EOF Then
      Exit Do
   End If
   Do While sMov = rsCta!cMovNro
      sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " " & rsCta!dDocFecha & " "
      Linea sTexto, oImpresora.gPrnCondensadaON & " " & Space(23) & sDoc & "  " & oImpresora.gPrnCondensadaOFF, , nLin
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   If rsCta.EOF Then
      Exit Do
   End If
Loop

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & Space(87) & "TOTAL DIARIO  : " & PrnVal(nDebeD, 16, 2, False) & "  " & PrnVal(nHaberD, 16, 2, False) & "  " & PrnVal(nSaldo, 16, 2) & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnBoldOFF, 2
nLin = nLin + 2
N = N + 1
Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, oImpresora.gPrnCondensadaON & " -------------------------" & String(36, "-") & "------------------------------------------------------------------------------------------------"
Linea sTexto, oImpresora.gPrnBoldON & Space(87) & "TOTAL CUENTA  : " & PrnVal(nDebe, 16, 2, False) & "  " & PrnVal(nHaber, 16, 2, False) & oImpresora.gPrnBoldOFF, 2
nLin = nLin + 2

Linea sTexto, ImprimeMayorCuentaCab(nLin, P, pdFechaDel, pdFechaAl, psCtaCod, psCtaDes, pnLinPage), 0
Linea sTexto, Space(87) & "SALDO INICIAL : " & PrnVal(nSaldoIni, 16, 2) & Space(5) & "SALDO INICIAL ME : " & PrnVal(pnSaldoIniME, 16, 2) ' Se Agrego Saldo Inicial ME GITU
Linea sTexto, Space(87) & "SALDO FINAL   : " & PrnVal(nSaldo, 16, 2) & Space(5) & "SALDO FINAL ME   : " & PrnVal(pnSaldoFinalME, 16, 2)  ' Se Agrego Saldo Final ME GITU
Linea sTexto, oImpresora.gPrnCondensadaOFF & oImpresora.gPrnBoldOFF
nLin = pnLinPage + 1
Linea sTexto, oImpresora.gPrnCondensadaON & " =========================" & String(36, "=") & "================================================================================================" & oImpresora.gPrnCondensadaOFF
ImprimeMayorCta = sTexto
RaiseEvent BarraClose
rs.Close: Set rs = Nothing
rsCta.Close: Set rsCta = Nothing

End Function

Private Function ImprimeMayorCuentaCab(ByRef nLin As Integer, ByRef P As Integer, dFechaDel As Date, dFechaAl As Date, sCtaCod As String, sCtaDes As String, pnLinPage As Integer) As String
Dim K As Integer, nObjs As Integer
Dim sCabe As String
Dim oConec As New COMConecta.DCOMConecta

If nLin > pnLinPage - 6 Then
   oConec.AbreConexion
   Linea sCabe, oImpresora.gPrnBoldON & Cabecera(" M A Y O R   D E   C U E N T A S    C O N T A B L E S", P, "", 80, , Format(oConec.GetFechaHoraServer(), gsFormatoFechaView)), 0
   oConec.CierraConexion
   Linea sCabe, oImpresora.gPrnBoldON & Centra("Del " & dFechaDel & " al " & dFechaAl, 80), 2
   Linea sCabe, oImpresora.gPrnCondensadaON & " ======================" & String(36, "=") & "===================================================================================================="
   Linea sCabe, "  Número               " & "     DOCUMENTO                       " & "                                                          MOVIMIENTO                         "
   Linea sCabe, "  de                   " & " ------------------------------------" & " DESCRIPCION                                     --------------------------        S A L D O "
   Linea sCabe, "  Movimiento           " & " Tipo  Número              Fecha     " & "                                                       DEBE          HABER                   "
   Linea sCabe, " ----------------------" & String(36, "-") & "----------------------------------------------------------------------------------------------------" & oImpresora.gPrnCondensadaOFF
   Linea sCabe, oImpresora.gPrnBoldON & "CUENTA CONTABLE : " & Justifica(sCtaCod, 16) & " " & sCtaDes & oImpresora.gPrnBoldOFF, 2
   nLin = 12
End If
ImprimeMayorCuentaCab = sCabe
Set oConec = Nothing
End Function

Public Function ImprimeFlujoCtasBancos(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpeCod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As String

Dim N As Integer
Dim P As Integer
Dim nLin As Integer
Dim lnTotD As Currency, lnTotH As Currency, lnSaldo As Currency
Dim lbOk As Boolean
Dim lsCond As String
Dim lsMovNro As String
Dim lsDocNro As String
Dim lsCtaIFCod As String
Dim lnImporte As Currency
Dim lsTexto As String
Dim lsGlosa As String
Dim rs As ADODB.Recordset
Dim lsPersCod As String
Dim lsSimbolo As String

Dim oCont As NCajaCtaIF
Set oCont = New NCajaCtaIF
Set rs = New ADODB.Recordset
On Error GoTo ErrRep
nLin = pnLinPage
lsTexto = ""
RaiseEvent BarraShow(1)
RaiseEvent BarraProgress(0, "Generando Flujo de Bancos", "Cargando Datos...", "", vbBlue)

lsSimbolo = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "S/.", "US$")
Set rs = oCont.GetFlujosCtaIF(psOpeCod, psCtaIfCodDesde, psCtaIfCodHasta, pdDesde, pdHasta)
RaiseEvent BarraProgress(1, "Generando Flujo de Bancos", "Cargando Datos...", "", vbBlue)
RaiseEvent BarraClose

If rs.EOF And rs.BOF Then
    rs.Close: Set rs = Nothing
    ImprimeFlujoCtasBancos = ""
    Exit Function
End If
RaiseEvent BarraShow(rs.RecordCount)
lsCtaIFCod = ""
lnImporte = 0
lnTotD = 0
lnTotH = 0
Do While Not rs.EOF
    If nLin >= 56 Then
        nLin = pnLinPage
    End If
    CabeceraFlujo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdDesde, pdHasta
    If rs!cMovNro = lsMovNro And IIf(Mid(psOpeCod, 3, 1) = gMonedaExtranjera, rs!nMovMEImporte, rs!nMovImporte) = lnImporte Then
       Linea lsTexto, oImpresora.gPrnCondensadaON & Space(64) & ImpreFormat(rs!Documento, 20) & " " & oImpresora.gPrnCondensadaOFF
    Else
       If rs!cPerscod + rs!cCtaIFCod <> lsPersCod + lsCtaIFCod Then
           If lnTotD > 0 Or lnTotH > 0 Then
               Linea lsTexto, oImpresora.gPrnCondensadaON & " ------------------------------------------------------------------------------------------------------------------------------------"
               nLin = nLin + 1
               Linea lsTexto, oImpresora.gPrnBoldON & Space(75) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, 2
               nLin = nLin + 2
           End If
           lnTotD = 0
           lnTotH = 0
           Linea lsTexto, oImpresora.gPrnBoldON & oImpresora.gPrnCondensadaON & " CUENTA : [" & rs!cCtaIFCod & "] " & rs!cNomCtaIF & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnBoldOFF
           nLin = nLin + 1
           lnTotD = 0: lnTotH = 0: lnSaldo = 0
           Linea lsTexto, oImpresora.gPrnCondensadaON & " SALDO ANTERIOR..." & Space(100) & " " & Right(Space(14) & Format(IIf(Mid(psOpeCod, 3, 1) = gMonedaExtranjera, rs!nSaldoIniME, rs!nSaldoIni), "#,#0.00"), 13) & oImpresora.gPrnCondensadaOFF
           If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
                lnSaldo = rs!nSaldoIniME
           Else
                lnSaldo = rs!nSaldoIni
           End If
           nLin = nLin + 1
       End If
       lsMovNro = rs!cMovNro
       If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
           lnImporte = rs!nMovMEImporte
           lnTotD = lnTotD + rs!DebeME
           lnTotH = lnTotH + rs!HaberME
       Else
           lnImporte = rs!nMovImporte
           lnTotD = lnTotD + rs!Debe
           lnTotH = lnTotH + rs!Haber
       End If
       lnSaldo = lnSaldo + lnImporte
       lsGlosa = rs!cMovDesc
       
       Linea lsTexto, oImpresora.gPrnCondensadaON & " " & Mid(lsMovNro, 1, 8) & "-" & Mid(lsMovNro, 9, 6) & "-" & Right(lsMovNro, 4) & ImpreFormat(Replace(lsGlosa, Chr(13), ""), 40) & " " _
               & ImpreFormat(rs!Documento, 17) & " " _
               & ImpreFormat(IIf(Mid(psOpeCod, 3, 1) = gMonedaExtranjera, rs!DebeME, rs!Debe), 14, 2, True) & ImpreFormat(IIf(Mid(psOpeCod, 3, 1) = gMonedaExtranjera, rs!HaberME, rs!Haber), 14, 2, True) & " " _
               & ImpreFormat(lnSaldo, 10, 2, True) & oImpresora.gPrnCondensadaOFF
   
       nLin = nLin + 1
    End If
    RaiseEvent BarraProgress(rs.Bookmark, "Generando Flujo de Bancos", Trim(rs!cpersNombre) & " [" & rs!cCtaIFCod + "]", "", vbBlue)
    lsCtaIFCod = rs!cCtaIFCod
    lsPersCod = rs!cPerscod
    rs.MoveNext
Loop
If lnTotD > 0 Or lnTotH > 0 Then
    Linea lsTexto, oImpresora.gPrnCondensadaON & " ------------------------------------------------------------------------------------------------------------------------------------"
    nLin = nLin + 1
    Linea lsTexto, oImpresora.gPrnBoldON & Space(75) & "TOTALES  " & ImpreFormat(lnTotD, 14, 2, True) & ImpreFormat(lnTotH, 14, 2, True) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, 2
    nLin = nLin + 2
End If
Linea lsTexto, oImpresora.gPrnCondensadaON & " ====================================================================================================================================" & oImpresora.gPrnCondensadaOFF, 2
ImprimeFlujoCtasBancos = lsTexto
rs.Close
RaiseEvent BarraClose
Set rs = Nothing
Exit Function
ErrRep:
    ImprimeFlujoCtasBancos = ""
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function CabeceraFlujo(ByRef lsTexto As String, ByRef nLin As Integer, ByRef P As Integer, _
            ByVal pnLinPage As Integer, pnColPage As Integer, ByVal psSimbolo As String, ByVal pdDesde As Date, ByVal pdHasta As Date) As String
If nLin >= pnLinPage - 7 Then
   Linea lsTexto, Cabecera(" FLUJO  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage, , , gsNomCmac)
   Linea lsTexto, Centra("( DEL " & pdDesde & " AL " & pdHasta & ")", pnColPage) & oImpresora.gPrnNegritaOFF
   Linea lsTexto, oImpresora.gPrnCondensadaON
   nLin = 11
   Linea lsTexto, " ===================================================================================================================================="
   Linea lsTexto, " Número de             CONCEPTO                                       DOCUMENTO                        MOVIMIENTO             SALDOS "
   Linea lsTexto, " Movimiento                                                          Tipo  Número                 DEBE           HABER               "
   Linea lsTexto, " ------------------------------------------------------------------------------------------------------------------------------------" & PrnSet("B-")
End If
End Function

Public Function ImprimeSaldosCtaIf(ByVal pnColPage As Integer, ByVal pnLinPage As Integer, ByVal psOpeCod As String, _
                                ByVal pdFecha As Date) As String
Dim rs As ADODB.Recordset
Dim nLin  As Integer
Dim oCtaIf As COMNAuditoria.NCajaCtaIF
Dim lsTexto As String
Dim lsSimbolo As String
Dim P As Integer
Set oCtaIf = New COMNAuditoria.NCajaCtaIF
Set rs = New ADODB.Recordset
lsTexto = ""
lsSimbolo = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "S/.", "$.")
Set rs = oCtaIf.GetRepSaldosCtaIf(psOpeCod, pdFecha, , gEstadoCtaIFActiva, Mid(psOpeCod, 3, 1))
nLin = pnLinPage
If Not rs.EOF And Not rs.BOF Then
    RaiseEvent BarraShow(rs.RecordCount)
    Do While Not rs.EOF
       If Len(rs!cCtaIFCod) = 15 Then
          Linea lsTexto, ""
          nLin = nLin + 1
       End If
       RaiseEvent BarraProgress(rs.Bookmark, "Generando Saldos de Cuentas", "", "", vbBlue)
       CabeceraRepoSaldo lsTexto, nLin, P, pnLinPage, pnColPage, lsSimbolo, pdFecha, psOpeCod
       Linea lsTexto, CON & " " & Mid(rs!cCtaIFCod & Space(20), 1, 25) & "   " & Mid(Trim(rs!cBancoDesc) & " " & rs!cCtaIFDesc & Space(70), 1, 55) & " " _
              & Right(Space(18) & Format(rs!Debe, "#,#0.00"), 18) & " " & Right(Space(10) & Format(rs!Haber, "#,#0.00"), 18) & " " & COFF
       nLin = nLin + 1
       rs.MoveNext
    Loop
    Linea lsTexto, CON & " ====================================================================================================================================" & COFF, 2
    RaiseEvent BarraClose
End If
rs.Close: Set rs = Nothing
ImprimeSaldosCtaIf = lsTexto
End Function

Private Sub CabeceraRepoSaldo(ByRef lsTexto As String, ByRef nLin As Integer, ByRef P As Integer, _
                        ByVal pnLinPage As Integer, pnColPage As Integer, ByVal psSimbolo As String, ByVal pdHasta As Date, ByVal psOpeCod As String)
If nLin > pnLinPage - 6 Then
   Linea lsTexto, Cabecera(" SALDOS  DE  ENTIDADES  FINANCIERAS ", P, psSimbolo, pnColPage, , , gsNomCmac)
   Linea lsTexto, Centra("( AL " & pdHasta & " )", pnColPage) & BOFF
   Linea lsTexto, CON
   nLin = 11
   Linea lsTexto, " ============================================================================================================================"
   Linea lsTexto, "   ENTIDAD FINANCIERA                                                                                 S A L D O S   M. " & IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "N.", "E.")
   Linea lsTexto, "                                                                                                    DEBE              HABER  "
   Linea lsTexto, " ----------------------------------------------------------------------------------------------------------------------------"
End If
End Sub

Public Function ImprimeSaldoProvision(psSaldoProv As String, pssaldoCanc As String, pssaldo As String, lsfechadel As String, lsfechaal As String) As String
Dim sTexto As String
Dim nRow As Integer
Dim sMov As String
Dim sDoc As String
Dim N As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberH As Currency
Dim nHaberD As Currency
Dim nSaldo  As Currency, nSaldoIni As Currency
Dim sFecha  As String
Dim rs      As ADODB.Recordset
Dim rsCta   As ADODB.Recordset
Dim nLin    As Integer
Dim P       As Integer
Dim sTipoCta As String

nLin = 1
P = 0
Dim oCont As New NContAsientos
Dim oSdo  As New NCtasaldo
Set oSdo = Nothing

sTexto = ""

Linea sTexto, CON & BON & Space(20) & "  SALDO DE PROVISION DE PAGO A PROVEEDORES DEL DIA : " & lsfechadel & "  AL  :   " & lsfechaal & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(10) & "----------------------------------------------------------------------------" & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(20) & "Saldo Provisionado         Saldo Cancelado         Saldo  " & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(10) & "----------------------------------------------------------------------------" & COFF & BOFF, 2, nLin
Linea sTexto, CON & BON & Space(20) & psSaldoProv & "                    " & pssaldoCanc & "                " & pssaldo & "          " & COFF & BOFF, 2, nLin
ImprimeSaldoProvision = sTexto

End Function

Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "") 'DAOR 20070816, Se aumentó subtitulo
    Dim EspHorxPag As Integer
    EspHorxPag = 170
    psCadImp = psCadImp & oImpresora.gPrnSaltoLinea
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "FECHA :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(76) & "FECHA   :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    End If
End Sub

Private Sub ImprimeCabeceraDocumentoAncho(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "") 'DAOR 20070816, Se aumentó subtitulo
    Dim EspHorxPag As Integer
    EspHorxPag = 280
    psCadImp = psCadImp & oImpresora.gPrnSaltoLinea
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(190) & "FECHA :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(169) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(76) & "FECHA   :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    End If
End Sub

Private Sub ImprimeTotalesCarteraDetxIntDSP(psCadImp As String, ByVal pnTotAprobado As Double, _
                                    ByVal pnTotSaldoK As Double, ByVal pnTotDevengado As Double, _
                                    ByVal pnTotSuspenso As Double, ByVal pnTotProvision As Double, ByVal pnCuantos As Integer, ByVal psGrupo As String)

        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & String(90, "-") & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & " MONTO APRO.           SALDO K       DEVENGADO      SUSPENSO         PROVISION " & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(79)
        psCadImp = psCadImp & String(90, "-") & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & psGrupo & pnCuantos & " Caso(s)" & oImpresora.gPrnSaltoLinea
        psCadImp = psCadImp & Space(82)
        
        psCadImp = psCadImp & ImpreFormat(Format(pnTotAprobado, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotSaldoK, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotDevengado, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotSuspenso, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & Space(3)
        psCadImp = psCadImp & ImpreFormat(Format(pnTotProvision, "#,###,###.##"), 12, 2)
        psCadImp = psCadImp & oImpresora.gPrnSaltoLinea
        pnTotAprobado = 0
        pnTotSaldoK = 0
        pnTotDevengado = 0
        pnTotSuspenso = 0
        pnTotProvision = 0
End Sub
