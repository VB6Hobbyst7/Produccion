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
  'psTitulo = psTitulo
  lsCabe02 = psCodRepo & String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
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
            Optional ByVal pnValortasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0) As String

Dim lsCadImp As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35

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
        pnValortasacion = lrCredPig!nTasacion
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
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
Else
    Do While Not prPersonas.EOF
        If Right(lsNombre, 3) <> " / " Then
            lsCodPers = prPersonas!cPersCod
            lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
            lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
            lsDireccion = IIf(IsNull(prPersonas!cPersDireccDomicilio), "", prPersonas!cPersDireccDomicilio)
            lszona = Mid(Trim(prPersonas!Zona), 1, 25)
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
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lszona) & Chr(10)  ' Zona
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsCiudad) & Chr(10)  ' Ciudad
    lsCadImp = lsCadImp & Space(32) & ImpreCarEsp(lsTelefono) & Chr(10)  ' Telefono
    lsCadImp = lsCadImp & Space(32) & " " & Chr(10)  ' Telefono
    
    
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnPiezas, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnOroBruto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnOroNeto, 9, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(72) & ImpreFormat(pnValortasacion, 9, 2, True) & Chr(10)
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

' Imprime el Contrato Pignoraticio con Detalle de Joyas
Public Function nPrintContratoPignoraticioDetANTxx(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValortasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0) As String

Dim lsCadImp As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
'*******************
'Agregado MPBR
Dim lsApellido As String
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
Dim lmDetalle(20) As String
Dim pnInteresDif As Double, pnMontoRec As Double


If pbBuscaData = True Then

    Dim lrCredPig As ADODB.Recordset
    Dim lrCredPigJoyas As ADODB.Recordset
    Dim lrCredPigPersonas As ADODB.Recordset
    Dim lrCredPigCostos As ADODB.Recordset
    Dim lrCredPigJoyasDet As ADODB.Recordset
    
    Dim loMuestraContrato As DColPContrato
    
    Set loMuestraContrato = New DColPContrato
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
        Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCodCta)
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
        pnValortasacion = lrCredPig!nTasacion
        pnMontoPrestado = lrCredPig!nMontoCol
        pnInteresDif = IIf(IsNull(lrCredPig!nIntDif), 0, lrCredPig!nIntDif)
        pnMontoRec = pnMontoPrestado - pnInteresDif
         pnPlazo = IIf(IsNull(lrCredPig!nPlazo) = True, 0, lrCredPig!nPlazo)
        psFecContrato = Format(lrCredPig!dVigencia, "mm/dd/yyyy")
        psFecVencimiento = Format(DateAdd("d", pnPlazo, lrCredPig!dVigencia), "mm/dd/yyyy")
        psLote = lrCredPig!cLote
        'psLote = fgEliminaEnters(lrCredPig!cLote)
        
       
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
        Dim item As Integer, itemTMP As Integer
        item = 0
        Do While Not lrCredPigJoyasDet.EOF
            lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                        & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
            item = item + 1
            lrCredPigJoyasDet.MoveNext
        Loop
        For itemTMP = item To 20
            lmDetalle(itemTMP) = ""
        Next
        lrCredPigJoyasDet.Close
        
    End If
 
End If

lnEspacio = 7

lsDescLote = psLote

lsCadImp = ""


lsDescLote = lsDescLote & Chr$(13)
'********************************************
'Modificacion MPBR
'Select Case prPersonas.RecordCount
'    Case 1
'        lnLongi = 35
'    Case 2
'        lnLongi = 25
'    Case 3
'        lnLongi = 18
'    Case 4
'        lnLongi = 12
'End Select
        
prPersonas.MoveFirst
'If lnLongi = 35 Then
Dim sSql As String, rstemp As ADODB.Recordset, cn As DConecta

 Set rstemp = New ADODB.Recordset
 Set cn = New DConecta
  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre "
  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
  If rstemp.State = 1 Then rstemp.Close
  cn.AbreConexion
  'rsTemp.Open ssql, cn.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
  Set rstemp = cn.CargaRecordSet(sSql)
  If rstemp.State = 1 Then
    If rstemp.RecordCount > 0 Then
       lsApellido = rstemp!cpersapellido
       lsNombre = rstemp!cPersNombre
    End If
    If rstemp.State = 1 Then rstemp.Close
  End If
  cn.CierraConexion
  Set cn = Nothing

    lsCodPers = prPersonas!cPersCod
'    lsApellido = prPersonas!cpersapellido
'    lsNombre = prPersonas!cpersnombre
    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
'Else
'    Do While Not prPersonas.EOF
'        If Right(lsNombre, 3) <> " / " Then
'            lsCodPers = prPersonas!cPersCod
'            lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
'            lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
'            lsDireccion = IIf(IsNull(prPersonas!cPersDireccDomicilio), "", prPersonas!cPersDireccDomicilio)
'            lszona = Mid(Trim(prPersonas!Zona), 1, 25)
'            lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
'            lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
'        End If
'        lsNombre = lsNombre & Left(prPersonas!cpersnombre, lnLongi) & " / "
'        prPersonas.MoveNext
'    Loop
'    lsNombre = Left(lsNombre, Len(lsNombre) - 3)
'End If

    lsCadImp = ""
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(103)              'Tamaño  : 80, 77, 103
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)
'    '--lsCadImp = lsCadImp & Chr$(27) & Chr$(77)
'    '---lsCadImp = lsCadImp & Chr$(27) & Chr$(120) & Chr$(2)
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'tipo de letra courier
'
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(20)
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(1)
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(65) & Chr$(9)     'Espaciado
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(9)      'Espaciado
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(2)
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(30) '--???? sevolo 10012005
'
'    '--1lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'tamaño pag
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(18)
    
'************
    
    
  '  lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
  '  lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
  '  lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
  '  lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
  '  lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
  '  lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
  '  lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita

    lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    'lsCadImp = lsCadImp & " " & String( 12, Chr(10))
        
   ' lsCadImp = lsCadImp & " " & String(11, Chr(10))
   ' lsCadImp = lsCadImp & " " & String(11, Chr(10))
    lsCadImp = lsCadImp & Space(10) & lsCodPers & Space(25) & Format(psFecContrato, "dd/mm/yyyy") & Chr(10) '12
    lsCadImp = lsCadImp & Space(45) & Format(psFecVencimiento, "dd/mm/yyyy")            '13
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
    lsCadImp = lsCadImp & Space(25) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) '& "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita  75
    
    lsCadImp = lsCadImp & Chr(10) & Chr(10)  '15 &chr(10)
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsApellido)       '16
    lsCadImp = lsCadImp & Space(80 - Len(lsApellido)) & ImpreCarEsp(Str(pnPlazo)) & Chr(10)
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsNombre) & Chr(10) & Chr(10)      '16
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsNroDocIdentidad) & " " & ImpreCarEsp(lsNroDocTributario) & Chr(10)   ' Doc identid   19
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsDireccion) & " " & Chr(10)   ' Direccion 20
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lszona) & Chr(10) & Chr(10)  ' Zona    22
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsCiudad) & Chr(10)  ' Ciudad  23
    lsCadImp = lsCadImp & Space(15) & ImpreCarEsp(lsTelefono) & Chr(10) & Chr(10) ' Telefono   24
'    lsCadImp = lsCadImp & Space(16) & " " '& Chr(10)  ' Telefono
        
    'lsCadImp = lsCadImp & Chr(10) & Char(10)

    lsCadImp = lsCadImp & Chr$(27) & Chr$(65) & Chr$(12)     'Espaciado

    lsCadImp = lsCadImp & Space(70) & ImpreFormat(pnPiezas, 9, 2, True) & Chr(10)                   '27
    lsCadImp = lsCadImp & Space(70) & ImpreFormat(pnOroBruto, 9, 2, True) & Chr(10)                 '28
    lsCadImp = lsCadImp & Space(70) & ImpreFormat(pnOroNeto, 9, 2, True) & Chr(10)         '30
    lsCadImp = lsCadImp & Space(70) & ImpreFormat(pnValortasacion, 9, 2, True) & Chr(10)   '31
    lsCadImp = lsCadImp & Space(70) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(70) & ImpreFormat(pnMontoPrestado, 9, 2, True) & Chr(10)   '33
    'lsCadImp = lsCadImp & Space(67) & ImpreFormat(pnInteresDif, 9, 2, True) & Chr(10)
    'lsCadImp = lsCadImp & Space(67) & ImpreFormat(pnMontoRec, 9, 2, True) & Chr(10)
    
    lsCadImp = lsCadImp & Space(3) & "Item" & Space(3) & "Descripción" & Space(22) _
                & "P. Neto" & Space(21) & ImpreFormat(pnInteresDif, 9, 2, True)

    
    lsCadImp = lsCadImp & Space(70) & "" & Chr(10)
                     
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(0), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(1), 55) & Space(2) & ImpreFormat(pnCostoTasacion, 18, 2, True) & Chr(10) '" " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(2), 55) & Space(2) & ImpreFormat(pnCostoCustodia, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(3), 55) & Space(2) & ImpreFormat(pnImpuesto, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(4), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(5), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(6), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(7), 55) & ImpreFormat(pnMontoRec, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(8), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(9), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(10), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(11), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(12), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(13), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(14), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(15), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(16), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(17), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(18), 55) & " " & Chr(10)
    lsCadImp = lsCadImp & Space(3) & ImpreFormat(lmDetalle(19), 55) & " " & Chr(10)
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
   ' lsCadImp = lsCadImp & Chr(10)
    '
    '***Modificar poner a la izquierda
    lsCadImp = lsCadImp & Space(10) & "INC. 11: Pasados 30 dias del" & Chr(10)
    lsCadImp = lsCadImp & Space(10) & " vencimiento  de su Contrato" & Chr(10)
    lsCadImp = lsCadImp & Space(10) & ImpreCarEsp(" sus Joyas entrarán a Remate") & Chr(10)
    lsCadImp = lsCadImp & Space(10) & " sin Notificar." & Chr(10)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita
    'lsCadImp = lsCadImp & Space(15) & "" & Chr(10) & Chr(10)
    'lsCadImp = lsCadImp & Space(13) & "TASA EFECTIVA ANUAL: " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
    If pnNroDuplicado > 0 Then
        lsCadImp = lsCadImp & Space(13) & "Duplicado Nro : " & ImpreFormat(pnNroDuplicado, 5, 0) & Chr(10)
    End If
    'lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    'lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(13) & "" & Chr(10) & Chr(10) '& Chr(10) & Chr(10) & Chr(10)

    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita

    lsCadImp = lsCadImp & Chr(10) & Chr(10) & Space(20) & psCodUser
    lsCadImp = lsCadImp & Space(40) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) '& "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Space(10) & "" & Chr(10)
    lsCadImp = lsCadImp & Space(70) & Str(pnPlazo)
    'lsCadImp = lsCadImp & Chr(27) + "0"
    'La mejor Opcion
   ' lsCadImp = lsCadImp & Chr$(27) & Chr$(77) 'Chr$(27) + Chr$(107) + Chr$(0)
    lsCadImp = lsCadImp & Chr(10)
    ' lsCadImp = lsCadImp & " " & String(11, Chr(10))
    
    'Configuracion Original
    lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
    lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita

    
    nPrintContratoPignoraticioDetANTxx = lsCadImp

End Function

' Imprime el Contrato Pignoraticio con Detalle de Joyas
Public Function nPrintContratoPignoraticioDet(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValortasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0) As String


'Original
Dim lsCadImp As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
'*******************
'Agregado MPBR
Dim lsApellido As String
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
Dim lmDetalle(20) As String
Dim pnInteresDif As Double, pnMontoRec As Double

Dim cTempoTexto75 As String * 75

If pbBuscaData = True Then

    Dim lrCredPig As ADODB.Recordset
    Dim lrCredPigJoyas As ADODB.Recordset
    Dim lrCredPigPersonas As ADODB.Recordset
    Dim lrCredPigCostos As ADODB.Recordset
    Dim lrCredPigJoyasDet As ADODB.Recordset
    
    Dim loMuestraContrato As DColPContrato
    
    Set loMuestraContrato = New DColPContrato
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
        Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCodCta)
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
        pnValortasacion = lrCredPig!nTasacion
        pnMontoPrestado = lrCredPig!nMontoCol
        pnInteresDif = IIf(IsNull(lrCredPig!nIntDif), pnInteres, lrCredPig!nIntDif)
        pnMontoRec = pnMontoPrestado - pnInteresDif
         pnPlazo = IIf(IsNull(lrCredPig!nPlazo) = True, 0, lrCredPig!nPlazo)
        psFecContrato = Format(lrCredPig!dVigencia, "dd/MM/yyyy")
        psFecVencimiento = Format(DateAdd("d", pnPlazo, lrCredPig!dVigencia), "dd/MM/yyyy")
        psLote = lrCredPig!cLote
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
        Set lrCredPigPersonas = Nothing
        
        'Costos de Credito
        pnInteres = lrCredPigCostos!nInteres
        pnCostoTasacion = lrCredPigCostos!nTasacion
        pnCostoCustodia = lrCredPigCostos!nCustodia
        pnImpuesto = lrCredPigCostos!nImpuesto
        
        lrCredPigCostos.Close
        Set lrCredPigCostos = Nothing
        Dim item As Integer, itemTMP As Integer
        item = 0
        Do While Not lrCredPigJoyasDet.EOF
            lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                        & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
            item = item + 1
            lrCredPigJoyasDet.MoveNext
        Loop
        For itemTMP = item To 20
            lmDetalle(itemTMP) = ""
        Next
        lrCredPigJoyasDet.Close
        
    End If
 
End If

lnEspacio = 7
lsDescLote = psLote

lsCadImp = ""

lsDescLote = lsDescLote & Chr$(13)
        
prPersonas.MoveFirst

Dim sSql As String, rstemp As ADODB.Recordset, cn As DConecta

 Set rstemp = New ADODB.Recordset
 Set cn = New DConecta
  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre "
  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
  If rstemp.State = 1 Then rstemp.Close
  cn.AbreConexion
  Set rstemp = cn.CargaRecordSet(sSql)
  If rstemp.State = 1 Then
    If rstemp.RecordCount > 0 Then
       lsApellido = rstemp!cpersapellido
       lsNombre = rstemp!cPersNombre
    End If
    If rstemp.State = 1 Then rstemp.Close
  End If
  cn.CierraConexion
  Set cn = Nothing

    lsCodPers = prPersonas!cPersCod
    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
    lsTelefono = prPersonas!cPersTelefono

'Fin Original



'Dim loPrevio As New Previo.clsPrevio
'Dim lsCadImp As String
'Dim i As Integer

'i = 0

    lsCadImp = ""
    lsCadImp = Chr$(27) & Chr$(64)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) 'Condensada
    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(103)
    lsCadImp = lsCadImp & "   " & Chr(10)
    lsCadImp = lsCadImp & "   " & Chr(10)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)
    
    '****MODIFICACION PARA QUE NO SALGA CON NEGRITA
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'activa negrita
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
     
    'If i > 0 Then
    '    lsCadImp = lsCadImp & "" & Chr(10)
    '    lsCadImp = lsCadImp & "" & Chr(10)
    '    lsCadImp = lsCadImp & "" & Chr(10)
    '    lsCadImp = lsCadImp & "" & Chr(10)
    'End If
     
     
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & Space(28) & lsCodPers & Space(42) & Format(psFecContrato, "dd/MM/yyyy") & Space(38) & psCodCta & Chr(10)
    lsCadImp = lsCadImp & Space(83) & Format(psFecVencimiento, "dd/MM/yyyy") & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsApellido) & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsNombre) & Chr(10)
    
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreCarEsp(lsNroDocIdentidad) & " " & ImpreCarEsp(lsNroDocTributario)
    
    lsCadImp = lsCadImp & Space(38) & cTempoTexto75 & Space(24) & ImpreCarEsp(Str(pnPlazo)) & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsDireccion) & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lszona) & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsCiudad) & Chr(10)
    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsTelefono) & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnPiezas, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnOroBruto, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnOroNeto, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnValortasacion, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    
    'Titulo y Monto Prestamo
    cTempoTexto75 = ""
    cTempoTexto75 = "Item" & Space(3) & "Descripción" & Space(22) & "P. Neto"    '& Space(21) & ImpreFormat(pnInteresDif, 9, 2, True)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnMontoPrestado, 18, 2, True) & Chr(10)
    
    'Primer Item y el Interes
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(0), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnInteresDif, 18, 2, True) & Chr(10)
    
    'Segundo Item no mas
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(1), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    
    'tercer Item y Costo de Tasacion
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(2), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnCostoTasacion, 18, 2, True) & Chr(10)
    
    'Cuarto Item y Costo de Custodia
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(3), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnCostoCustodia, 18, 2, True) & Chr(10)
    
    'Quinto Item y el Impuesto
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(4), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnImpuesto, 18, 2, True) & Chr(10)
    
    'Item 6
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(5), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    
    'Item 7
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(6), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    
    'Item 8
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(7), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    
    'Item 9 y Neto a Recibir
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(8), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(18) '*****QUITA CONDENSADA
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'activa negrita
    lsCadImp = lsCadImp & ImpreFormat(pnMontoRec, 18, 2, True) & Chr(10)
    lsCadImp = lsCadImp & Chr$(27) + Chr$(70) ' desactiva negrita
    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) '*****PONE CONDENSADA
    
    'Item 10 a 20
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(9), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(10), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(11), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(12), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(13), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(14), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(15), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(16), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(17), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(18), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    cTempoTexto75 = ""
    cTempoTexto75 = ImpreFormat(lmDetalle(19), 55)
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
    
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & Space(30) & "Inc. 11: pasados 30 dias del vencimiento de su contrato sus joyas entrarán a Remate Sin Notificar" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & Space(116) & psCodCta & Chr(10)
    
    cTempoTexto75 = ""
    If pnNroDuplicado > 0 Then
        cTempoTexto75 = "Duplicado Nro : " & ImpreFormat(pnNroDuplicado, 5, 0)
    End If
    
    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(1) & psCodUser & Space(25) & ImpreCarEsp(Str(pnPlazo)) & Chr(10)
    
    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada

    nPrintContratoPignoraticioDet = lsCadImp
        
End Function



'Public Function nPrintContratoPignoraticioDetMon(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
'            Optional ByVal prPersonas As Recordset, _
'            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
'            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
'            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
'            Optional ByVal pnValortasacion As Currency, _
'            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
'            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
'            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
'            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0) As String
'
'
''Original
'Dim lsCadImp As String
'
'Dim lsDescLote As String
'
'Dim lnEspacio As Integer
'Dim lnLongi As Integer, lnX As Integer
''*******************
''Agregado MPBR
'Dim lsApellido As String
'Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
'Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
'Dim lmDetalle(20) As String
'Dim pnInteresDif As Double, pnMontoRec As Double
'
'Dim cTempoTexto75 As String * 75
'
'If pbBuscaData = True Then
'
'    Dim lrCredPig As ADODB.Recordset
'    Dim lrCredPigJoyas As ADODB.Recordset
'    Dim lrCredPigPersonas As ADODB.Recordset
'    Dim lrCredPigCostos As ADODB.Recordset
'    Dim lrCredPigJoyasDet As ADODB.Recordset
'
'    Dim loMuestraContrato As DColPContrato
'
'    Set loMuestraContrato = New DColPContrato
'        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
'        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
'        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
'        Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
'        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCodCta)
'    Set loMuestraContrato = Nothing
'
'    If lrCredPig.BOF And lrCredPig.EOF Then
'        Set lrCredPig = Nothing
'        Set lrCredPigJoyas = Nothing
'        Set lrCredPigPersonas = Nothing
'        MsgBox " No se encuentra el Credito Pignoraticio " & psCodCta, vbInformation, " Aviso "
'        Exit Function
'    Else
'        pnOroBruto = lrCredPig!nOroBruto
'        pnOroNeto = lrCredPig!noroneto
'        pnPiezas = lrCredPig!npiezas
'        pnValortasacion = lrCredPig!nTasacion
'        pnMontoPrestado = lrCredPig!nMontoCol
'        pnInteresDif = IIf(IsNull(lrCredPig!nIntDif), pnInteres, lrCredPig!nIntDif)
'        pnMontoRec = pnMontoPrestado - pnInteresDif
'         pnPlazo = IIf(IsNull(lrCredPig!nPlazo) = True, 0, lrCredPig!nPlazo)
'        psFecContrato = Format(lrCredPig!dVigencia, "dd/MM/yyyy")
'        psFecVencimiento = Format(DateAdd("d", pnPlazo, lrCredPig!dVigencia), "dd/MM/yyyy")
'        psLote = lrCredPig!cLote
'        pnTasaInteresMensual = lrCredPig!nTasaInteres
'        lrCredPig.Close
'        Set lrCredPig = Nothing
'
'        ' Kilatajes
'        pn14K = lrCredPigJoyas!nK14
'        pn16K = lrCredPigJoyas!nK16
'        pn18K = lrCredPigJoyas!nK18
'        pn21K = lrCredPigJoyas!nK21
'
'        lrCredPigJoyas.Close
'        Set lrCredPigJoyas = Nothing
'
'        ' Mostrar Clientes
'        Set prPersonas = lrCredPigPersonas
'        Set lrCredPigPersonas = Nothing
'
'        'Costos de Credito
'        pnInteres = lrCredPigCostos!nInteres
'        pnCostoTasacion = lrCredPigCostos!nTasacion
'        pnCostoCustodia = lrCredPigCostos!nCustodia
'        pnImpuesto = lrCredPigCostos!nImpuesto
'
'        lrCredPigCostos.Close
'        Set lrCredPigCostos = Nothing
'        Dim item As Integer, itemTMP As Integer
'        item = 0
'        Do While Not lrCredPigJoyasDet.EOF
'            lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
'                        & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
'            item = item + 1
'            lrCredPigJoyasDet.MoveNext
'        Loop
'        For itemTMP = item To 20
'            lmDetalle(itemTMP) = ""
'        Next
'        lrCredPigJoyasDet.Close
'
'    End If
'
'End If
'
'lnEspacio = 7
'lsDescLote = psLote
'
'lsCadImp = ""
'
'lsDescLote = lsDescLote & Chr$(13)
'
'prPersonas.MoveFirst
'
'Dim sSql As String, rstemp As ADODB.Recordset, cn As DConecta
'
' Set rstemp = New ADODB.Recordset
' Set cn = New DConecta
'  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
'  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre "
'  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
'  If rstemp.State = 1 Then rstemp.Close
'  cn.AbreConexion
'  Set rstemp = cn.CargaRecordSet(sSql)
'  If rstemp.State = 1 Then
'    If rstemp.RecordCount > 0 Then
'       lsApellido = rstemp!cpersapellido
'       lsNombre = rstemp!cPersNombre
'    End If
'    If rstemp.State = 1 Then rstemp.Close
'  End If
'  cn.CierraConexion
'  Set cn = Nothing
'
'    lsCodPers = prPersonas!cPersCod
'    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
'    lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
'    lsDireccion = prPersonas!cPersDireccDomicilio
'    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
'    lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
'    lsTelefono = prPersonas!cPersTelefono
'
''Fin Original
'
'
'
''Dim loPrevio As New Previo.clsPrevio
''Dim lsCadImp As String
''Dim i As Integer
'
''i = 0
'
'    lsCadImp = ""
'
'
'
'    lsCadImp = Chr$(27) & Chr$(64)
'
'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) 'Condensada
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
'
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
'
'    '*lsCadImp = lsCadImp & Chr$(27) & Chr$(103) 'tam letra 15cpi
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
'    lsCadImp = lsCadImp & "   " & Chr(10)
'    lsCadImp = lsCadImp & "   " & Chr(10)
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77) ' tam letra 10 cpi
'
'
'    '****MODIFICACION PARA QUE NO SALGA CON NEGRITA
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'activa negrita
'    'lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
'
'    'If i > 0 Then
'    '    lsCadImp = lsCadImp & "" & Chr(10)
'    '    lsCadImp = lsCadImp & "" & Chr(10)
'    '    lsCadImp = lsCadImp & "" & Chr(10)
'    '    lsCadImp = lsCadImp & "" & Chr(10)
'    'End If
'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) 'Condensada
'
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & Space(28) & lsCodPers & Space(42) & Format(psFecContrato, "dd/MM/yyyy") & Space(38) & psCodCta & Chr(10)
'    lsCadImp = lsCadImp & Space(83) & Format(psFecVencimiento, "dd/MM/yyyy") & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsApellido) & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsNombre) & Chr(10)
'
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreCarEsp(lsNroDocIdentidad) & " " & ImpreCarEsp(lsNroDocTributario)
'
'    lsCadImp = lsCadImp & Space(38) & cTempoTexto75 & Space(24) & ImpreCarEsp(Str(pnPlazo)) & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsDireccion) & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lszona) & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsCiudad) & Chr(10)
'    lsCadImp = lsCadImp & Space(38) & ImpreCarEsp(lsTelefono) & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnPiezas, 18, 2, True) & Chr(10)
'    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnOroBruto, 18, 2, True) & Chr(10)
'    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnOroNeto, 18, 2, True) & Chr(10)
'    lsCadImp = lsCadImp & Space(38 + 79) & ImpreFormat(pnValortasacion, 18, 2, True) & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'
'    'Titulo y Monto Prestamo
'    cTempoTexto75 = ""
'    cTempoTexto75 = "Item" & Space(3) & "Descripción" & Space(22) & "P. Neto"    '& Space(21) & ImpreFormat(pnInteresDif, 9, 2, True)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnMontoPrestado, 18, 2, True) & Chr(10)
'
'    'Primer Item y el Interes
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(0), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnInteresDif, 18, 2, True) & Chr(10)
'
'    'Segundo Item no mas
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(1), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'
'    'tercer Item y Costo de Tasacion
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(2), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnCostoTasacion, 18, 2, True) & Chr(10)
'
'    'Cuarto Item y Costo de Custodia
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(3), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnCostoCustodia, 18, 2, True) & Chr(10)
'
'    'Quinto Item y el Impuesto
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(4), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22) & ImpreFormat(pnImpuesto, 18, 2, True) & Chr(10)
'
'    'Item 6
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(5), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'
'    'Item 7
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(6), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'
'    'Item 8
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(7), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'
'    'Item 9 y Neto a Recibir
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(8), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(22)
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(18) '*****QUITA CONDENSADA
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'activa negrita
'    lsCadImp = lsCadImp & ImpreFormat(pnMontoRec, 18, 2, True) & Chr(10)
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(70) ' desactiva negrita
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) '*****PONE CONDENSADA
'
'    'Item 10 a 20
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(9), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(10), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(11), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(12), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(13), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(14), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(15), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(16), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(17), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(18), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'    cTempoTexto75 = ""
'    cTempoTexto75 = ImpreFormat(lmDetalle(19), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Chr(10)
'
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & Space(30) & "Inc. 11: pasados 30 dias del vencimiento de su contrato sus joyas entrarán a Remate Sin Notificar" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & "" & Chr(10)
'    lsCadImp = lsCadImp & Space(116) & psCodCta & Chr(10)
'
'    cTempoTexto75 = ""
'    If pnNroDuplicado > 0 Then
'        cTempoTexto75 = "Duplicado Nro : " & ImpreFormat(pnNroDuplicado, 5, 0)
'    End If
'
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(1) & psCodUser & Space(25) & ImpreCarEsp(Str(pnPlazo)) & Chr(10)
'
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
'
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(77)
'
'
'
'    nPrintContratoPignoraticioDet = lsCadImp
'
'End Function


'Procedimiento de impresión del recibo de renovación
Public Function nPrintReciboRenovacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnSaldoCapAnt As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnInteresMor As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, ByVal psVarPersCMAC As String, _
        ByVal psGastoCorrespondencia As Currency, Optional ByVal psMensaje As String = "", Optional ByVal pnITF As Double = 0#) As String
Dim lsCadImp As String
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
    
    
    Dim vEspacio As Integer
    vEspacio = 6
        lsCadImp = "" & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Renovacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(1) & Space(10 + vEspacio) _
                            & "Pignoraticio-Renovacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & "" & ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio) _
                                 & "" & ImpreFormat(psVarPersCMAC, 41, 0) & Chr(10)
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & "Contrato  : " & lsCodAntiguo & Space(10) & Space(1 + vEspacio) _
                                & "Contrato  : " & lsCodAntiguo & Space(10) & Chr(10)
        Else
            lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                                & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        End If
        'ImpreFormat(psNombCliente, 28)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Chr(10)
        
        lsCadImp = lsCadImp & "Saldo          :" & ImpreFormat(pnSaldoCapAnt, 21, , True) & Space(vEspacio + 2) _
                            & "Saldo          :" & ImpreFormat(pnSaldoCapAnt, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 2) _
                            & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Chr(10)
        
        lsCadImp = lsCadImp & "Mora           :" & ImpreFormat(pnInteresMor, 21, , True) & Space(vEspacio + 2) _
                            & "Mora           :" & ImpreFormat(pnInteresMor, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 2) _
                            & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ImpreFormat(pnCostoCustodia + pnImpuesto, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Custodia :" & ImpreFormat(pnCostoCustodia + pnImpuesto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Notif.   :" & ImpreFormat(psGastoCorrespondencia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Notif.   :" & ImpreFormat(psGastoCorrespondencia, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Chr(10)
        
        lsCadImp = lsCadImp & "NUEVO SALDO    :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio + 2) _
                            & "NUEVO SALDO    :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Renovación Nro : ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio + 2) _
                            & "" & ImpreCarEsp("Renovación Nro : ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Chr(10)
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
                            
    nPrintReciboRenovacion = lsCadImp
End Function

'Procedimiento de impresión del recibo de cancelacion del contrato
Public Function nPrintReciboCancelacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnMontoPrestamo As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal psCodUser As String, _
        ByVal psVarPersCMAC As String, Optional ByVal psMensaje As String = "", Optional psGastoCorrespondencia As Currency, Optional ByVal pnITF As Double = 0#) As String
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 6
        lsCadImp = "" & Chr(10) '& Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Cancelacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(10 + vEspacio) _
                            & "Pignoraticio-Cancelacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        If psVarPersCMAC <> "" Then
            lsCadImp = lsCadImp & "" & ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio) _
                                & "" & ImpreFormat(psVarPersCMAC, 41, 0) & Chr(10)
        End If
        'ImpreFormat(psNombCliente, 28)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Chr(10)
        
        lsCadImp = lsCadImp & "Prestamo       :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Space(vEspacio + 2) _
                            & "Prestamo       :" & ImpreFormat(pnMontoPrestamo, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 2) _
                            & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 2) _
                            & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Impuesto       :" & ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio + 2) _
                            & "Impuesto       :" & ImpreFormat(pnImpuesto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Custodia :" & ImpreFormat(pnCostoCustodia, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Chr(10)
                                
        lsCadImp = lsCadImp & "Costo Notif.   :" & ImpreFormat(psGastoCorrespondencia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Notif.   :" & ImpreFormat(psGastoCorrespondencia, 21, , True) & Chr(10)
                                
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & "SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio) _
                            & "SALDO DE CAPITAL :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & "" & ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Chr(10)
        'lsCadImp = lsCadImp & " " & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(13 + vEspacio) _
                            & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 38) & Space(vEspacio) _
                            & "" & ImpreFormat(psMensaje, 38) & Chr(10)
 
    nPrintReciboCancelacion = lsCadImp
End Function


Public Function nPrintReciboDesembolso(ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal pnMontoEntregar As Currency, _
        ByVal pnInteresCompensat As Currency, ByVal psNomAgencia As String, _
        ByVal psCodUser As String, ByVal pnITF As Double) As String
Dim lsCadImp As String
Dim psNombCliente As String
Dim Co As DConecta
Dim rs As New ADODB.Recordset
Dim sql As String
Set Co = New DConecta

sql = " select dbo.pstaNombre(cPersNombre,0) Nombre from Persona PE " & _
      " Inner Join ProductoPersona PP on PE.cPersCod = PP.cPersCod " & _
      " Where nPrdPersRelac =" & ColocRelacPers.gColRelPersTitular & " And cCtaCod ='" & psCtaCod & "'"
      
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
psNombCliente = rs!Nombre
Set rs = Nothing
Set Co = Nothing
    Dim vEspacio As Integer
    vEspacio = 6
    '--ImpreFormat(psNombCliente, 28)
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Desembolso" & Space(9) & Space(10 + vEspacio) _
                            & "Pignoraticio-Desembolso" & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & psFechaHora & Space(vEspacio) _
                            & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & psFechaHora & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & Space(1 + vEspacio) _
                            & "Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        lsCadImp = lsCadImp & "Fec.Prestamo: " & psFechaHora & "                  " & Space(vEspacio) _
                            & "Fec.Prestamo: " & psFechaHora & "                  " & Chr(10)
        
        lsCadImp = lsCadImp & "Prestamo       :" & ImpreFormat(pnSaldoCap, 21, , True) & Space(vEspacio + 2) _
                            & "Prestamo       :" & ImpreFormat(pnSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Interes Compen.:" & ImpreFormat(pnInteresCompensat, 21, , True) & Space(vEspacio + 2) _
                            & "Interes Compen.:" & ImpreFormat(pnInteresCompensat, 21, , True) & Chr(10)
                                                        
        lsCadImp = lsCadImp & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Space(vEspacio + 2) _
                            & "NETO A PAGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Chr(10) & Chr(10)
        'lsCadImp = lsCadImp & " MONTO ENTREGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Space(vEspacio) _
                            & " MONTO ENTREGAR   :" & ImpreFormat(pnMontoEntregar, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Desembolso  Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & "" & ImpreCarEsp("Desembolso  Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Chr(10)
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
        ByVal psMensaje As String, ByVal pnITF As Double) As String
    
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & Chr(10) & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Credito Pignoraticio " & Space(11) & Space(9 + vEspacio) _
                            & "Credito Pignoraticio " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio - 1) _
                            & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio - 1) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & "--------------------------------------" & Space(2 + vEspacio) & " --------------------------------------" & Chr(10)
        lsCadImp = lsCadImp & "Costo Duplicado  : " & ImpreFormat(pnCostoDuplicado, 15) & Space(5 + vEspacio) _
                            & "Costo Duplicado  : " & ImpreFormat(pnCostoDuplicado, 15) & Chr(10)
        lsCadImp = lsCadImp & "I.T.F.           : " & ImpreFormat(CDbl(pnITF), 15) & Space(5 + vEspacio) _
                            & "I.T.F.           : " & ImpreFormat(CDbl(pnITF), 15) & Chr(10)
        lsCadImp = lsCadImp & "Total Duplicado  : " & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(5 + vEspacio) _
                            & "Total Duplicado  : " & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15)
        lsCadImp = lsCadImp & Chr(10) & "" & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Space(4 + vEspacio) _
                            & "" & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & "" & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Duplicado Contrato ") & Space(23) & Space(vEspacio) _
                            & "" & ImpreCarEsp("Duplicado Contrato ") & Space(20) & Chr(10)
                            
        lsCadImp = lsCadImp & "--------------------------------------" & Space(5 + vEspacio) _
                            & "--------------------------------------" & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
        
        lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10)
        nPrintReciboDuplicadoContrato = lsCadImp
        
        
End Function

Public Function nPrintReciboCobroCustodia(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnCostoDuplicado As Currency, _
        ByVal pnNroDuplicado As Integer, pnTasaInteresMensual As Double, ByVal psCodUser As String, _
        ByVal psMensaje As String, ByVal pnITF As Double) As String
    
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & Chr(10) & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Credito Pignoraticio " & Space(11) & Space(9 + vEspacio) _
                            & "Credito Pignoraticio " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio - 1) _
                            & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio - 1) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & "--------------------------------------" & Space(2 + vEspacio) & " --------------------------------------" & Chr(10)
        lsCadImp = lsCadImp & "Cobro de Custodia:" & ImpreFormat(pnCostoDuplicado, 15) & Space(6 + vEspacio) _
                            & "Cobro de Custodia:" & ImpreFormat(pnCostoDuplicado, 15) & Chr(10)
        lsCadImp = lsCadImp & "I.T.F.           :" & ImpreFormat(CDbl(pnITF), 15) & Space(6 + vEspacio) _
                            & "I.T.F.           :" & ImpreFormat(CDbl(pnITF), 15) & Chr(10)
        lsCadImp = lsCadImp & "Total A Pagar    :" & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(6 + vEspacio) _
                            & "Total A Pagar    :" & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & "" & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & Chr(10)
                            
        lsCadImp = lsCadImp & "--------------------------------------" & Space(5 + vEspacio) _
                            & "--------------------------------------" & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
        
        lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10)
        nPrintReciboCobroCustodia = lsCadImp
        
        
End Function
' Impresión del recibo de Amortizacion
Public Function nPrintReciboAmortizacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnSaldoCap As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, ByVal psVarPersCMAC As String, Optional ByVal psMensaje As String = "", Optional ByVal pnITF As Double = 0#) As String
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 6
        lsCadImp = "" & Chr(10) & Chr(10)
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(15)  '*****activa condensada
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Amortizacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(9 + vEspacio) _
                            & "Pignoraticio-Amortizacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & "" & ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
                            
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio) _
                                 & " " & ImpreFormat(psVarPersCMAC, 41, 0) & Chr(10)
        End If
        'ImpreFormat(psNombCliente, 28)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & Chr(10)
        lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Chr(10)
        
        lsCadImp = lsCadImp & "Saldo          :" & ImpreFormat(pnSaldoCap, 21, , True) & Space(vEspacio + 2) _
                            & "Saldo          :" & ImpreFormat(pnSaldoCap, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 2) _
                            & "Capital        :" & ImpreFormat(pnCapital, 21, , True) & Chr(10)
        lsCadImp = lsCadImp & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 2) _
                            & "Interes        :" & ImpreFormat(pnInteres, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Impuesto       :" & ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio + 2) _
                            & "Impuesto       :" & ImpreFormat(pnImpuesto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Custodia :" & ImpreFormat(pnCostoCustodia, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Remate   :" & ImpreFormat(pnCostoRemate, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        'MODIFICACION MPBR
        lsCadImp = lsCadImp & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "I.T.F.         :" & ImpreFormat(pnITF, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "NETO A PAGAR   :" & ImpreFormat(pnMontoPago + pnITF, 21, , True) & Chr(10)
        
        
        'ANTES 31082004
        'lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio) _
        '                    & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 21, , True) & Chr(10)
                            
                            
        lsCadImp = lsCadImp & "NUEVO SALDO    :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio + 2) _
                            & "NUEVO SALDO    :" & ImpreFormat(pnNewSaldoCap, 21, , True) & Chr(10)
                            
        'lsCadImp = lsCadImp & " " & ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio) _
        '                    & " " & ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ImpreFormat(pnPlazo, 4, 0) & Chr(10)
        lsCadImp = lsCadImp & "" & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Chr(10)
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
                            
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(18)    '**para quitar condensada
                            
    nPrintReciboAmortizacion = lsCadImp
End Function

'Procedimiento de impresión del recibo de Extorno
Public Function nPrintReciboExtorno(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psOperacion As String, _
        ByVal pnMonto As Currency, ByVal pnSaldoCapAnt As Currency, ByVal pnMovExtornado As Long, _
        ByVal psCodUser As String, ByVal psVarPersCMAC As String, _
        Optional ByVal psMensaje As String = "", Optional ByVal psUsuOri As String) As String
Dim lsCadImp As String
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
    
    
    Dim vEspacio As Integer
    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(12) & Space(10 + vEspacio) _
                            & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 30, 0) & Space(12) & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 30, 0) & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Space(24 + vEspacio) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Chr(10)
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio) _
                                 & " " & ImpreFormat(psVarPersCMAC, 41, 0) & Chr(10)
        Else
            lsCadImp = lsCadImp & "" & Chr(10)
        End If
        lsCadImp = lsCadImp & "" & Chr(10)
        
        lsCadImp = lsCadImp & " ---- " & ImpreFormat("EXTORNO " & psOperacion, 25) & " ---- " & Space(4 + vEspacio) _
                            & " ---- " & ImpreFormat("EXTORNO " & psOperacion, 25) & " ---- " & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & Space(1 + vEspacio) _
                                & " Contrato  : " & lsCodAntiguo & Space(10) & Chr(10)
        Else
            lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                                & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Chr(10)
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 28) & Chr(10)
        lsCadImp = lsCadImp & "" & Chr(10)
        lsCadImp = lsCadImp & " MONTO     :" & ImpreFormat(pnMonto, 21, , True) & Space(7 + vEspacio) _
                            & " MONTO     :" & ImpreFormat(pnMonto, 21, , True) & Chr(10)
                            
        lsCadImp = lsCadImp & " -------------------------------------" & Space(5 + vEspacio) _
                            & " -------------------------------------" & Chr(10)
        lsCadImp = lsCadImp & Space(15) & "" & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Space(6 + vEspacio) _
                            & " " & ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Chr(10)
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Space(20 + vEspacio) & "  " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Space(20 + vEspacio) & "  " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Chr(10)
        
        lsCadImp = lsCadImp & Space(15) & "" & Chr(10)
        lsCadImp = lsCadImp & "" & Chr(10) & Chr(10)
        'lsCadImp = lsCadImp & "" & Chr(10) & Chr(10) & Chr(10)
    nPrintReciboExtorno = lsCadImp
End Function

Public Function nPrintReciboPagoSobrante(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnSobrante As Currency, _
        ByVal psCodUser As String, ByVal psMensaje As String) As String
    
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & Chr(10) & Chr(10) & Chr(10) & Chr(10)
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
        lsCadImp = lsCadImp & " CLIENTE:" & Trim(psNombCliente) & Space(vEspacio + 8) _
                            & " CLIENTE:" & Trim(psNombCliente) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & " --------------------------------------" & Space(5 + vEspacio) & " --------------------------------------" & Chr(10)
        lsCadImp = lsCadImp & " Sobrante         : " & ImpreFormat(pnSobrante, 15) & Space(5 + vEspacio) _
                            & " Sobrante         : " & ImpreFormat(pnSobrante, 15) & Chr(10)
        'lsCadImp = lsCadImp & " I.T.F.           : " & ImpreFormat(CDbl(pnITF), 15) & Space(5 + vEspacio) _
        '                    & " I.T.F.           : " & ImpreFormat(CDbl(pnITF), 15) & Chr(10)
        'lsCadImp = lsCadImp & " Total Duplicado  : " & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(5 + vEspacio) _
        '                    & " Total Duplicado  : " & ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Chr(10)
        'lsCadImp = lsCadImp & " " & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Space(4 + vEspacio) _
        '                    & " " & ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreCarEsp("Sobrante de Remate ") & Space(23) & Space(vEspacio) _
                            & " " & ImpreCarEsp("Sobrante de Remate ") & Space(20) & Chr(10)
                            
        lsCadImp = lsCadImp & " --------------------------------------" & Space(5 + vEspacio) _
                            & " --------------------------------------" & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
        
        lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10)
        nPrintReciboPagoSobrante = lsCadImp

            
End Function

