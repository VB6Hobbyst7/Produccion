VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColPImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Cabecera de Impresiones de Reportes de Pignoraticio
Public Function nImprimeCabeceraReportes(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psCodRepo As String) As String
        
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & Chr(10)
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & chr$(10)
  ' Cabecera 2
  lsCabe01 = lsCabe01 & FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)
  ' Titulo
  psTitulo = psCodRepo & " - " & psTitulo
  lsCabe02 = String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)
  ' Comenta
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportes = lsCabRepo
End Function

' Imprime el Contrato Pignoraticio
Public Function nPrintContratoPignoraticio(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0) As String

Dim lsCadImp As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lsZona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35

If pbBuscaData = True Then

    Dim lrCredPig As ADODB.Recordset
    Dim lrCredPigJoyas As ADODB.Recordset
    Dim lrCredPigPersonas As ADODB.Recordset
    Dim lrCredPigCostos As ADODB.Recordset
    
    Dim loMuestraContrato As DColPContrato
    
    Set loMuestraContrato = New DColPContrato
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
        Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
    Set loMuestraContrato = Nothing

    If lrCredPig.BOF And lrCredPig.EOF Then
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        Set lrCredPigPersonas = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psCodCta, vbInformation, " Aviso "
        Exit Function
    Else
        pnOroBruto = lrCredPig!nOroBruto
        pnOroNeto = lrCredPig!noroneto
        pnPiezas = lrCredPig!npiezas
        pnValorTasacion = lrCredPig!nTasacion
        pnMontoPrestado = lrCredPig!nMontoCol
        psFecContrato = Format(lrCredPig!dVigencia, "mm/dd/yyyy")
        psFecVencimiento = Format(lrCredPig!dFecVencIni, "mm/dd/yyyy")
        psLote = lrCredPig!cLote
        'psLote = fgEliminaEnters(lrCredPig!cLote)
        
        pnPlazo = lrCredPig!nPlazoIni
        pnTasaInteresMensual = lrCredPig!nTasaInteres

        lrCredPig.Close
        Set lrCredPig = Nothing

        ' Kilatajes
        pn14K = lrCredPigJoyas!nK14
        pn16K = lrCredPigJoyas!nK16
        pn18K = lrCredPigJoyas!nK18
        pn21K = lrCredPigJoyas!nK21
        
        lrCredPigJoyas.Close
        Set lrCredPigJoyas = Nothing

        ' Mostrar Clientes
        Set prPersonas = lrCredPigPersonas
        
        'lrCredPigPersonas.Close
        Set lrCredPigPersonas = Nothing
        
        'Costos de Credito
        pnInteres = lrCredPigCostos!nInteres
        pnCostoTasacion = lrCredPigCostos!nTasacion
        pnCostoCustodia = lrCredPigCostos!nCustodia
        pnImpuesto = lrCredPigCostos!nImpuesto
        
        lrCredPigCostos.Close
        Set lrCredPigCostos = Nothing
        
    End If
 
End If

lnEspacio = 7

lsDescLote = psLote

lsCadImp = ""

If pn14K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ImpreFormat(pn14K, 8, , False) & " grs. 14K."
If pn16K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ImpreFormat(pn16K, 8, , False) & " grs. 16K."
If pn18K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ImpreFormat(pn18K, 8, , False) & " grs. 18K."
If pn21K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ImpreFormat(pn21K, 8, , False) & " grs. 21K."

lsDescLote = lsDescLote & Chr$(13)
Select Case prPersonas.RecordCount
    Case 1
        lnLongi = 35
    Case 2
        lnLongi = 25
    Case 3
        lnLongi = 18
    Case 4
        lnLongi = 12
End Select
        
prPersonas.MoveFirst
If lnLongi = 35 Then
    lsCodPers = prPersonas!cPersCod
    lsNombre = Left(prPersonas!cPersNombre, lnLongi)
    lsNroDocIdentidad = prPersonas!NroDNI
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lsZona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
Else
    Do While Not prPersonas.EOF
        If Right(lsNombre, 3) <> " / " Then
            lsCodPers = prPersonas!cPersCod
            lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
            lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
            lsDireccion = IIf(IsNull(prPersonas!cPersDireccDomicilio), "", prPersonas!cPersDireccDomicilio)
            lsZona = Mid(Trim(prPersonas!Zona), 1, 25)
            lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
            lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
        End If
        lsNombre = lsNombre & Left(prPersonas!cPersNombre, lnLongi) & " / "
        prPersonas.MoveNext
    Loop
    lsNombre = Left(lsNombre, Len(lsNombre) - 3)
End If

    lsCadImp = " " & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    lsCadImp = lsCadImp & Space(23) & lsCodPers & Space(20) & Format(psFecContrato, "dd/mm/yyyy") & Chr(10)
    lsCadImp = lsCadImp & Space(56) & Format(psFecVencimiento, "dd/mm/yyyy")
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
    lsCadImp = lsCadImp & Space(7) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita  75

    lsCadImp = lsCadImp & Chr(10)
        
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsNombre) & Chr(10)
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsNroDocIdentidad) & " " & ImpreCarEsp(lsNroDocTributario) & Chr(10) ' Doc identid
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsDireccion) & " " & Chr(10)  ' Direccion
    lsCadImp = lsCadImp & Space(85) & ImpreCarEsp(Str(pnPlazo)) & Chr(10) 'Plazo
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsZona) & Chr(10)  ' Zona
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsCiudad) & Chr(10)  ' Ciudad
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsTelefono) & Chr(10)  ' Telefono
    lsCadImp = lsCadImp & Space(32) & " " & Chr(10)  ' Telefono
    
    
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnPiezas, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnOroBruto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnOroNeto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnValorTasacion, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnMontoPrestado, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnInteres, 9, 2, True) & Chr(10)
    
    lsCadImp = lsCadImp & Space(70) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 1) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(QuiebreTexto(lsDescLote, 1))) & ImpreFormat(pnCostoTasacion, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 2) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(QuiebreTexto(lsDescLote, 2))) & ImpreFormat(pnCostoCustodia, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 3) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(QuiebreTexto(lsDescLote, 3))) & ImpreFormat(pnImpuesto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 4) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 5) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 6) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 7) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(QuiebreTexto(lsDescLote, 7))) & ImpreFormat(pnMontoPrestado - pnInteres - pnCostoTasacion - pnCostoCustodia - pnImpuesto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 8) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 9) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 10) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 11) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 12) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 13) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 14) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 15) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 16) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 17) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(QuiebreTexto(lsDescLote, 17))) & "INC. 11: Pasados 30 dias del"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 18) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(QuiebreTexto(lsDescLote, 18))) & " vencimiento  de su Contrato"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 19) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(QuiebreTexto(lsDescLote, 19))) & ImpreCarEsp(" sus Joyas entrarán a Remate")
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & QuiebreTexto(lsDescLote, 20) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(QuiebreTexto(lsDescLote, 20))) & " sin Notificar."
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & "" & Chr(10) & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "TASA EFECTIVA ANUAL: " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
    If pnNroDuplicado > 0 Then
        lsCadImp = lsCadImp & Space(13) & "Duplicado Nro : " & ImpreFormat(pnNroDuplicado, 5, 0) & Chr(10)
    End If
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)


    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(45) & psCodUser
    lsCadImp = lsCadImp & Space(9) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Space(10) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(65) & Str(pnPlazo)

    nPrintContratoPignoraticio = lsCadImp

End Function


'Procedimiento de impresión del recibo de renovación
Public Function nPrintReciboRenovacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnMontoPrestamo As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, Optional ByVal psMensaje As String = "") As String
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Pignoraticio-Renovacion " & Space(8) & Space(10 + vEspacio) _
                            & " Pignoraticio-Renovacion " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                            & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Chr(10)
        lsCadImp = lsCadImp & " Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & " Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Chr(10)
        
        lsCadImp = lsCadImp & " Prestamo         :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Space(vEspacio) _
                            & " Prestamo         :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Capital          :" & ImpreFormat(pnCapital, 21, , True) & Space(vEspacio) _
                            & " Capital          :" & ImpreFormat(pnCapital, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & " Interes          :" & ImpreFormat(pnInteres, 21, , True) & Space(vEspacio) _
                            & " Interes          :" & ImpreFormat(pnInteres, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Impuesto         :" & ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio) _
                            & " Impuesto         :" & ImpreFormat(pnImpuesto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Costo Custodia   :" & ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio) _
                            & " Costo Custodia   :" & ImpreFormat(pnCostoCustodia, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Costo Remate     :" & ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio) _
                            & " Costo Remate     :" & ImpreFormat(pnCostoRemate, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio) _
                            & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & " SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio) _
                            & " SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio) _
                            & " " & ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 40) & Space(vEspacio) _
                            & " " & ImpreFormat(psMensaje, 40) & Chr(10)
                            
    nPrintReciboRenovacion = lsCadImp
End Function

'Procedimiento de impresión del recibo de cancelacion del contrato
Public Function nPrintReciboCancelacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnMontoPrestamo As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, _
        ByVal psCodUser As String, Optional ByVal psMensaje As String = "") As String
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Pignoraticio-Cancelacion" & Space(8) & Space(10 + vEspacio) _
                            & " Pignoraticio-Cancelacion" & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                            & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Chr(10)
        lsCadImp = lsCadImp & " Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & " Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Chr(10)
        
        lsCadImp = lsCadImp & " Prestamo         :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Space(vEspacio) _
                            & " Prestamo         :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Capital          :" & ImpreFormat(pnCapital, 21, , True) & Space(vEspacio) _
                            & " Capital          :" & ImpreFormat(pnCapital, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Interes          :" & ImpreFormat(pnInteres, 21, , True) & Space(vEspacio) _
                            & " Interes          :" & ImpreFormat(pnInteres, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Impuesto         :" & ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio) _
                            & " Impuesto         :" & ImpreFormat(pnImpuesto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Costo Custodia   :" & ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio) _
                            & " Costo Custodia   :" & ImpreFormat(pnCostoCustodia, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Costo Remate     :" & ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio) _
                            & " Costo Remate     :" & ImpreFormat(pnCostoRemate, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio) _
                            & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & " SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio) _
                            & " SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & " " & ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Chr(10)
        'lsCadImp = lsCadImp & " " & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(13 + vEspacio) _
                            & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 38) & Space(vEspacio) _
                            & " " & ImpreFormat(psMensaje, 38) & Chr(10)
 
    nPrintReciboCancelacion = lsCadImp
End Function


Public Function nPrintReciboDesembolso(ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal pnMontoEntregar As Currency, _
        ByVal pnInteresCompensat As Currency, ByVal psNomAgencia As String, ByVal psCodUser As String) As String
Dim lsCadImp As String
Dim psNombCliente As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Dim Sql As String
Set Co = New DConecta

Sql = " select dbo.pstaNombre(cPersNombre,0) Nombre from Persona PE " & _
      " Inner Join ProductoPersona PP on PE.cPersCod = PP.cPersCod " & _
      " Where nPrdPersRelac =" & ColocRelacPers.gColRelPersTitular & " And cCtaCod ='" & psCtaCod & "'"
      
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
psNombCliente = Rs!Nombre
Set Rs = Nothing
Set Co = Nothing
    Dim vEspacio As Integer
    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Pignoraticio-Desembolso" & Space(8) & Space(10 + vEspacio) _
                            & " Pignoraticio-Desembolso" & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFechaHora, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFechaHora, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & Space(1 + vEspacio) _
                            & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Chr(10)
        lsCadImp = lsCadImp & " Fec.Prestamo: " & Format(psFechaHora, "dd/mm/yyyy") & "                  " & Space(vEspacio) _
                            & " Fec.Prestamo: " & Format(psFechaHora, "dd/mm/yyyy") & "                  " & Chr(10)
        
        lsCadImp = lsCadImp & " Prestamo         :" & ImpreFormat(pnSaldoCap, 21, , True) & Space(vEspacio) _
                            & " Prestamo         :" & ImpreFormat(pnSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " Interes Compen.  :" & ImpreFormat(pnInteresCompensat, 21, , True) & Space(vEspacio) _
                            & " Interes Compen.  :" & ImpreFormat(pnInteresCompensat, 21, , True) & Chr(10)
                                                        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoEntregar, 21, , True) & Space(vEspacio) _
                            & " NETO A PAGAR     :" & ImpreFormat(pnMontoEntregar, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & " MONTO ENTREGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Space(vEspacio) _
                            & " MONTO ENTREGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Desembolso  Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & " " & ImpreCarEsp("Desembolso  Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Chr(10)
        'lsCadImp = lsCadImp & " " & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & chr(10)
                            
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        'lsCadImp = lsCadImp & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(13 + vEspacio) _
        '                    & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
                            
        'lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 38) & Space(vEspacio) _
        '                    & " " & ImpreFormat(psMensaje, 38) & Chr(10)
 
    nPrintReciboDesembolso = lsCadImp
End Function





Public Function nPrintReciboDuplicadoContrato(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnCostoDuplicado As Currency, _
        ByVal pnNroDuplicado As Integer, pnTasaInteresMensual As Double, ByVal psCodUser As String, _
        ByVal psMensaje As String) As String
    
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Credito Pignoraticio " & Space(11) & Space(10 + vEspacio) _
                            & " Credito Pignoraticio " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio) _
                            & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Costo Duplicado  : " & ImpreFormat(pnCostoDuplicado, 20) & Space(vEspacio) _
                            & " Costo Duplicado  : " & ImpreFormat(pnCostoDuplicado, 20) & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & " " & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Duplicado Contrato ") & Space(23) & Space(vEspacio) _
                            & " " & ImpreCarEsp("Duplicado Contrato ") & Space(20) & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
        
        nPrintReciboDuplicadoContrato = lsCadImp
        
 
End Function

