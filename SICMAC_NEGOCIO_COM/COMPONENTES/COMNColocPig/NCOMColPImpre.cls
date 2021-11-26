VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gcPEN As New COMDConstantes.DCOMConstantes 'MARG ERS044-2016
'Cabecera de Impresiones de Reportes de Pignoraticio
Public Function nImprimeCabeceraReportes(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psCodRepo As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  Dim ofun As New COMFunciones.FCOMImpresion
  Dim oFunI As New COMFunciones.FCOMVarImpresion
  oFunI.Inicia psImpresora
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = ofun.FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & oFunI.gPrnSaltoLinea
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & oFunI.gPrnSaltoLinea
  ' Cabecera 2
  lsCabe01 = lsCabe01 & ofun.FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
  ' Titulo
  'psTitulo = psTitulo
  lsCabe02 = psCodRepo & String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & oFunI.gPrnSaltoLinea
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & oFunI.gPrnSaltoLinea
  ' Comenta
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & oFunI.gPrnSaltoLinea, "")
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportes = lsCabRepo
  Set oFunI = Nothing
End Function

' Imprime el Contrato Pignoraticio
Public Function nPrintContratoPignoraticio(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As ADODB.Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON) As String

Dim lsCadImp As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

If pbBuscaData = True Then

    Dim lrCredPig As New ADODB.Recordset
    Dim lrCredPigJoyas As New ADODB.Recordset
    Dim lrCredPigPersonas As New ADODB.Recordset
    Dim lrCredPigCostos As New ADODB.Recordset
    
    Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato
    
  
    Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
    Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
    Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
    Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
    Set loMuestraContrato = Nothing

    If lrCredPig.BOF And lrCredPig.EOF Then
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        Set lrCredPigPersonas = Nothing
        psmensaje = " No se encuentra el Credito Pignoraticio " & psCodCta
        Exit Function
    Else
        pnOroBruto = lrCredPig!nOroBruto
        pnOroNeto = lrCredPig!nOroNeto
        pnPiezas = lrCredPig!nPiezas
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

If pn14K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ofun.ImpreFormat(pn14K, 8, , False) & " grs. 14K."
If pn16K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ofun.ImpreFormat(pn16K, 8, , False) & " grs. 16K."
If pn18K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ofun.ImpreFormat(pn18K, 8, , False) & " grs. 18K."
If pn21K > 0 Then lsDescLote = lsDescLote & Chr$(13) & ofun.ImpreFormat(pn21K, 8, , False) & " grs. 21K."

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
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
Else
    Do While Not prPersonas.EOF
        If Right(lsNombre, 3) <> " / " Then
            lsCodPers = prPersonas!cPersCod
            lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
            lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
            lsDireccion = IIf(IsNull(prPersonas!cPersDireccDomicilio), "", prPersonas!cPersDireccDomicilio)
            lszona = Mid(Trim(prPersonas!Zona), 1, 25)
            lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
            lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
        End If
        lsNombre = lsNombre & Left(prPersonas!cPersNombre, lnLongi) & " / "
        prPersonas.MoveNext
    Loop
    lsNombre = Left(lsNombre, Len(lsNombre) - 3)
End If

    lsCadImp = " " & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(23) & lsCodPers & Space(20) & Format(psFecContrato, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(56) & Format(psFecVencimiento, "dd/mm/yyyy")
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
    lsCadImp = lsCadImp & Space(7) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) '& "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita  75

    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lsNombre) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea  ' Doc identid
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lsDireccion) & " " & oFunI.gPrnSaltoLinea   ' Direccion
    lsCadImp = lsCadImp & Space(85) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea  'Plazo
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea   ' Zona
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea   ' Ciudad
    lsCadImp = lsCadImp & Space(32) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea   ' Telefono
    lsCadImp = lsCadImp & Space(32) & " " & oFunI.gPrnSaltoLinea   ' Telefono
    
    
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnPiezas, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnOroBruto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnOroNeto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnValorTasacion, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnMontoPrestado, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnInteres, 9, 2, True) & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & Space(70) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 1) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(ofun.QuiebreTexto(lsDescLote, 1))) & ofun.ImpreFormat(pnCostoTasacion, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 2) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(ofun.QuiebreTexto(lsDescLote, 2))) & ofun.ImpreFormat(pnCostoCustodia, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 3) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(ofun.QuiebreTexto(lsDescLote, 3))) & ofun.ImpreFormat(pnImpuesto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 4) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 5) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 6) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 7) & " "
    lsCadImp = lsCadImp & Space(72 - 15 - Len(ofun.QuiebreTexto(lsDescLote, 7))) & ofun.ImpreFormat(pnMontoPrestado - pnInteres - pnCostoTasacion - pnCostoCustodia - pnImpuesto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 8) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 9) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 10) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 11) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 12) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 13) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 14) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 15) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 16) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 17) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 17))) & "INC. 11: Pasados 30 dias del"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 18) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 18))) & " vencimiento  de su Contrato"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea     'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 19) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 19))) & ofun.ImpreCarEsp(" sus Joyas entrarán a Remate")
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & ofun.QuiebreTexto(lsDescLote, 20) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 20))) & " sin Notificar."
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "TASA EFECTIVA ANUAL: " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & oFunI.gPrnSaltoLinea
    If pnNroDuplicado > 0 Then
        lsCadImp = lsCadImp & Space(13) & "Duplicado Nro : " & ofun.ImpreFormat(pnNroDuplicado, 5, 0) & oFunI.gPrnSaltoLinea
    End If
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea


    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(45) & psCodUser
    lsCadImp = lsCadImp & Space(9) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) ' & "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Space(10) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(65) & Str(pnPlazo)

    nPrintContratoPignoraticio = lsCadImp
    Set oFunI = Nothing
End Function

' Imprime el Contrato Pignoraticio con Detalle de Joyas
Public Function nPrintContratoPignoraticioDetANTxx(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As ADODB.Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim ofun As New COMFunciones.FCOMImpresion

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

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


If pbBuscaData = True Then

    Dim lrCredPig As New ADODB.Recordset
    Dim lrCredPigJoyas As New ADODB.Recordset
    Dim lrCredPigPersonas As New ADODB.Recordset
    Dim lrCredPigCostos As New ADODB.Recordset
    Dim lrCredPigJoyasDet As New ADODB.Recordset
    
    Dim loMuestraContrato As DCOMColPContrato
    
    Set loMuestraContrato = New DCOMColPContrato
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
        psmensaje = " No se encuentra el Credito Pignoraticio " & psCodCta
        Exit Function
    Else
        pnOroBruto = lrCredPig!nOroBruto
        pnOroNeto = lrCredPig!nOroNeto
        pnPiezas = lrCredPig!nPiezas
        pnValorTasacion = lrCredPig!nTasacion
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
            lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                        & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
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

Dim sSql As String, RSTEMP As ADODB.Recordset, cn As COMConecta.DCOMConecta

 Set RSTEMP = New ADODB.Recordset
 Set cn = New COMConecta.DCOMConecta
  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre "
  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
  If RSTEMP.State = 1 Then RSTEMP.Close
  cn.AbreConexion
  Set RSTEMP = cn.CargaRecordSet(sSql)
  If RSTEMP.State = 1 Then
    If RSTEMP.RecordCount > 0 Then
       lsApellido = RSTEMP!cPersApellido
       lsNombre = RSTEMP!cPersNombre
    End If
    If RSTEMP.State = 1 Then RSTEMP.Close
  End If
  cn.CierraConexion
  Set cn = Nothing

    lsCodPers = prPersonas!cPersCod
'    lsApellido = prPersonas!cpersapellido
'    lsNombre = prPersonas!cpersnombre
    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
   lsCadImp = lsCadImp & Space(10) & lsCodPers & Space(25) & Format(psFecContrato, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea  '12
    lsCadImp = lsCadImp & Space(45) & Format(psFecVencimiento, "dd/mm/yyyy")            '13
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
    lsCadImp = lsCadImp & Space(25) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) '& "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita  75
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea    '15 &oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsApellido)       '16
    lsCadImp = lsCadImp & Space(80 - Len(lsApellido)) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsNombre) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea        '16
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea    ' Doc identid   19
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsDireccion) & " " & oFunI.gPrnSaltoLinea    ' Direccion 20
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea    ' Zona    22
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea   ' Ciudad  23
    lsCadImp = lsCadImp & Space(15) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea   ' Telefono   24

    lsCadImp = lsCadImp & Chr$(27) & Chr$(65) & oFunI.gPrnSaltoPagina      'Espaciado

    lsCadImp = lsCadImp & Space(70) & ofun.ImpreFormat(pnPiezas, 9, 2, True) & oFunI.gPrnSaltoLinea                   '27
    lsCadImp = lsCadImp & Space(70) & ofun.ImpreFormat(pnOroBruto, 9, 2, True) & oFunI.gPrnSaltoLinea                 '28
    lsCadImp = lsCadImp & Space(70) & ofun.ImpreFormat(pnOroNeto, 9, 2, True) & oFunI.gPrnSaltoLinea         '30
    lsCadImp = lsCadImp & Space(70) & ofun.ImpreFormat(pnValorTasacion, 9, 2, True) & oFunI.gPrnSaltoLinea   '31
    lsCadImp = lsCadImp & Space(70) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(70) & ofun.ImpreFormat(pnMontoPrestado, 9, 2, True) & oFunI.gPrnSaltoLinea   '33
    
    lsCadImp = lsCadImp & Space(3) & "Item" & Space(3) & "Descripción" & Space(22) _
                & "P. Neto" & Space(21) & ofun.ImpreFormat(pnInteresDif, 9, 2, True)

    
    lsCadImp = lsCadImp & Space(70) & "" & oFunI.gPrnSaltoLinea
                     
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(0), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(1), 55) & Space(2) & ofun.ImpreFormat(pnCostoTasacion, 18, 2, True) & oFunI.gPrnSaltoLinea  '" " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(2), 55) & Space(2) & ofun.ImpreFormat(pnCostoCustodia, 18, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(3), 55) & Space(2) & ofun.ImpreFormat(pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(4), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(5), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(6), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(7), 55) & ofun.ImpreFormat(pnMontoRec, 18, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(8), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(9), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(10), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(11), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(12), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(13), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(14), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(15), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(16), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(17), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(18), 55) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(lmDetalle(19), 55) & " " & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita

    '***Modificar poner a la izquierda
    lsCadImp = lsCadImp & Space(10) & "INC. 11: Pasados 30 dias del" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(10) & " vencimiento  de su Contrato" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(10) & ofun.ImpreCarEsp(" sus Joyas entrarán a Remate") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(10) & " sin Notificar." & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    'lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(13) & "TASA EFECTIVA ANUAL: " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & oFunI.gPrnSaltoLinea
    If pnNroDuplicado > 0 Then
        lsCadImp = lsCadImp & Space(13) & "Duplicado Nro : " & ofun.ImpreFormat(pnNroDuplicado, 5, 0) & oFunI.gPrnSaltoLinea
    End If
    'lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea   '& oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita

    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Space(20) & psCodUser
    lsCadImp = lsCadImp & Space(40) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) '& "-" & InicialBoveda(Mid(psCodCta, 1, 5))
    lsCadImp = lsCadImp & Space(10) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(70) & Str(pnPlazo)
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
 
    'Configuracion Original
    lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
    lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita

    
    nPrintContratoPignoraticioDetANTxx = lsCadImp
    Set oFunI = Nothing
End Function

'**DAOR 20070115, Imprimir Volante de Tasación
Public Function nPrintVolanteTasacion(pnPlazo As Integer, pnOroBruto As Double, pnOroNeto As Double, pnValorTasacion As Currency, pnPiezas As Integer, prjoyas As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim lnLineas As Integer, lnCont As Integer, i As Integer
    oFunI.Inicia psImpresora
    lsCadImp = ""
    
    '**Configuración de página**********************************************************
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(66)   'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(18)  ' cancela condensada
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    '***********************************************************************************
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnPlazo, 3, 0) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Do While Not prjoyas.EOF
        lnCont = lnCont + 1
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Pzas"), 5, 0)
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Descripcion"), 33, 0)
        lsCadImp = lsCadImp & Space(4) & ofun.ImpreFormat(Right(prjoyas("Material"), 2), 10, 0)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PNeto"), 8, 2)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PBruto"), 8, 2) & oFunI.gPrnSaltoLinea
        If lnCont > 10 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(20) + "Continua siguiente pagina..."
            lnCont = 0
            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnPlazo, 3, 0) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        End If
        prjoyas.MoveNext
    Loop
    lnLineas = 14 - lnCont
    For i = 1 To lnLineas
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    Next
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(6) & ofun.ImpreFormat(pnPiezas, 5, 0) & Space(4) & ofun.ImpreFormat(pnOroBruto, 9, 2) & Space(25) & ofun.ImpreFormat(pnOroNeto, 9, 2) & Space(3) & ofun.ImpreFormat(pnValorTasacion, 9, 2) & oFunI.gPrnSaltoLinea
    nPrintVolanteTasacion = lsCadImp
    Set oFunI = Nothing
    Set ofun = Nothing
End Function

'**DAOR 20070117, Imprimir Hoja de Resumen
Public Function nPrintHojaResumen(psCodCta As String, Optional psNomCli As String, Optional psApeCli As String, Optional psNumDoc As String, Optional psDirCli As String, _
Optional ByVal pnValorTasacion As Currency, Optional ByVal pnMontoPrestado As Currency, _
Optional ByVal pnMontoRec As Currency, Optional ByVal pnITF As Double, _
Optional ByVal psFecContrato As Date, Optional ByVal psFecVencimiento As String, _
Optional ByVal pnInteres As Double, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal psmensaje As String, Optional ByRef HojaResumen1 As String, Optional ByRef HojaResumen2 As String) 'AMDO20130906 TI-ERS112-2013 As String

Dim lsCadImp As String, lsSM As String
Dim lnTasInt As Double, lnTasMor As Double, lnTicTEA As Double, lnTimTEA As Double, lnTimTEM As Double
Dim lnTasCus As Double, lnCosRei As Double, lnCosNtf As Double
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim lnLineas As Integer, lnCont As Integer, i As Integer
Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato
Dim lrCredPigTasas As New ADODB.Recordset
Dim loConstSis As New COMDConstSistema.NCOMConstSistema
Dim loParam As New COMDColocPig.DCOMColPCalculos
Dim loConstAge As New COMDConstantes.DCOMAgencias
Dim loConstZona As New COMDConstantes.DCOMZonas
Dim lrAgencia As New ADODB.Recordset
Dim lrZona As New ADODB.Recordset
Dim pnInteresDif As Double
Dim pnPlazo As Integer
'Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos


'*************************************************************--------------------
    Dim lrCredPig As New ADODB.Recordset
    Dim loMostraContrati As New COMDColocPig.DCOMColPContrato
   
    Set lrCredPig = loMostraContrati.dObtieneDatosCreditoPignoraticio(psCodCta)
    
    pnPlazo = IIf(IsNull(lrCredPig!nPlazo) = True, 0, lrCredPig!nPlazo)
    psFecVencimiento = Format(DateAdd("d", pnPlazo, lrCredPig!dVigencia), "dd/MM/yyyy")

'*************************************************************--------------------

    Set lrCredPigTasas = loMuestraContrato.dObtieneDatosCreditoPignoraticioTasas(psCodCta)
    If lrCredPigTasas.EOF Or lrCredPigTasas.BOF Then
        Set lrCredPigTasas = Nothing
        psmensaje = " No existen tasas de interes para esta cuenta" & psCodCta
        Exit Function
    Else
        While Not lrCredPigTasas.EOF
            Select Case lrCredPigTasas("nPrdTasaInteres")
             Case 1: lnTasInt = lrCredPigTasas("nTasaInteres")
             Case 3: lnTasMor = lrCredPigTasas("nTasaInteres")
            End Select
            lrCredPigTasas.MoveNext
        Wend
    End If
    
    lnTicTEA = Round((((1 + lnTasInt / 100) ^ 12 - 1) * 100), 2)
    lnTimTEA = Round((((1 + lnTasMor / 100) ^ 12 - 1) * 100), 2)
    '''lsSM = "S/. " 'marg ers044-2016
    lsSM = gcPEN.gcPEN_SIMBOLO & " " 'marg ers044-2016
    
    lnTimTEM = Round(lnTasMor, 2)
    lnCosRei = loParam.dObtieneColocParametro(gConsColPCostoDuplicadoContrato)
    'comentado por MADM 25/03/2010
    'lnCosNtf = loParam.dObtieneColocParametro(gConsColPCostoNotifAvisoVenci)
    If Mid(psCodCta, 4, 2) <> "" Then
'        Select Case CInt(Mid(psCodCta, 4, 2))
'            Case 1
'               lnCosNtf = loParam.dObtieneColocParametro(3103)
'            Case 2
'               lnCosNtf = loParam.dObtieneColocParametro(3104)
'            Case 3
'               lnCosNtf = loParam.dObtieneColocParametro(3105)
'            Case 4
'               lnCosNtf = loParam.dObtieneColocParametro(3106)
'            Case 5
'               lnCosNtf = loParam.dObtieneColocParametro(3107)
'            Case 6
'               lnCosNtf = loParam.dObtieneColocParametro(3108)
'            Case 7
'               lnCosNtf = loParam.dObtieneColocParametro(3109)
'            Case 9
'               lnCosNtf = loParam.dObtieneColocParametro(3111)
'            Case 10
'               lnCosNtf = loParam.dObtieneColocParametro(3112)
'            Case 12
'               lnCosNtf = loParam.dObtieneColocParametro(3113)
'            Case 13
'               lnCosNtf = loParam.dObtieneColocParametro(3114)
'            Case 24
'               lnCosNtf = loParam.dObtieneColocParametro(3115)
'            Case 25
'               lnCosNtf = loParam.dObtieneColocParametro(3116)
'            Case 31
'               lnCosNtf = loParam.dObtieneColocParametro(3117)
'        End Select
        'RECO20140722 ERS114-2014******************************************
        If Mid(psCodCta, 4, 2) = "33" Then
            Dim oColPNotif As New COMDColocPig.DCOMColPActualizaBD
            Dim drNotif As ADODB.Recordset
            
            Set oColPNotif = New COMDColocPig.DCOMColPActualizaBD
            Set drNotif = New ADODB.Recordset
            
            Set drNotif = oColPNotif.DevuelveValorNotificacionCarNotMinka(psCodCta)
            If Not (drNotif.EOF And drNotif.BOF) Then
                lnCosNtf = drNotif!nValor
            Else
                lnCosNtf = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", psCodCta)
            End If
        Else
            lnCosNtf = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", psCodCta)
        End If
        'RECO FIN**********************************************************
   End If
    '------------------------------------------------------------------------
     
    
    pnInteresDif = pnInteres
    Set lrAgencia = loConstAge.RecuperaAgencias(Mid$(psCodCta, 4, 2))
    Set lrZona = loConstZona.DameUnaZona(lrAgencia("cUbiGeoCod"))
        
    oFunI.Inicia psImpresora
    lsCadImp = ""
    
    '**Configuración de página**********************************************************
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(66)   'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(18)  ' cancela condensada
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    '***********************************************************************************
    
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 9)
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(pnMontoPrestado, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(pnInteres, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(pnMontoRec, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea AMDO20130906 TI-ERS112-2013
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(pnITF, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    '*** PEAC 20101220
'    lsCadImp = lsCadImp & Space(69) & ofun.ImpreFormat(lnTasInt, 12, 2) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(69) & ofun.ImpreFormat(lnTimTEM, 12, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(69) & ofun.ImpreFormat(lnTicTEA, 12, 2) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(69) & ofun.ImpreFormat(lnTimTEA, 12, 2) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    'FIN PEAC
    
    
    'lsCadImp = lsCadImp & Space(74) & ofun.ImpreFormat(Format(psFecVencimiento, "dd/MM/yyyy"), 10, 0) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(74) & Format(psFecVencimiento, "dd/MM/yyyy")
    'peac 20071219 Format(psFecVencimiento, "dd/MM/yyyy")
    
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 33)
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(lnCosNtf, 12, 2, True) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(65) & lsSM & ofun.ImpreFormat(lnCosRei, 12, 2, True)
'    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 11) AMDO20130906 TI-ERS112-2013
    
    'lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(lnTicTEA, 6, 0) & Space(33) & ofun.ImpreFormat(lnTimTEA, 6, 0) & oFunI.gPrnSaltoLinea
    
    '*** PEAC 20101220
    
    'lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(lnTicTEA, 6, 2) & Space(33) & ofun.ImpreFormat(lnTimTEA, 6, 2) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(30) & oFunI.gPrnSaltoLinea AMDO20130906 TI-ERS112-2013
    
    '*** FIN PEAC
    
    'AMDO 20130906 TI-ERS112-2013
    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
    HojaResumen1 = lsCadImp
    lsCadImp = ""
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    'END AMDO 20130906 TI-ERS112-2013
    
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 29) 'AMDO20130906
    lsCadImp = lsCadImp & Space(14) & ofun.ImpreFormat(lrZona("cUbigeoDescripcion"), 22, 0) & Day(psFecContrato) & Space(2) & ofun.ImpreFormat(Format(psFecContrato, "mmmm"), 15, 0) & Space(4) & Year(psFecContrato)
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 2)
    lsCadImp = lsCadImp & Space(20) & ofun.ImpreFormat(psNomCli & " " & psApeCli, 45, 0) & Space(5) & psNumDoc & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(20) & psApeCli & oFunI.gPrnSaltoLinea 'AMDO20130906
    lsCadImp = lsCadImp & Space(20) & ofun.ImpreFormat(psDirCli, 30, 0) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(23) & ofun.ImpreFormat(lrAgencia("cAgeDescripcion"), 20, 0) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Space(50) & ofun.ImpreFormat(lrAgencia("cAgeDireccion"), 30, 0)
    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
    
    'AMDO 20130906 TI-ERS112-2013
    HojaResumen2 = lsCadImp
    lsCadImp = ""
'    nPrintHojaResumen = lsCadImp
    'END AMDO 20130906 TI-ERS112-2013
    
    Set oFunI = Nothing
    Set ofun = Nothing
    Set loConstSis = Nothing
    Set loParam = Nothing
    Set loConstAge = Nothing
    Set loMuestraContrato = Nothing
    Set lrCredPigTasas = Nothing
    Set loConstZona = Nothing
    Set lrAgencia = Nothing
    Set lrZona = Nothing
End Function

' Imprime el Contrato Pignoraticio con Detalle de Joyas
Public Function nPrintContratoPignoraticioDet(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
            Optional ByVal prPersonas As ADODB.Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pnITF As Double) As String


'Original
Dim lsCadImp As String, lsSM As String

Dim lsDescLote As String

Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
'*******************
'Agregado MPBR
Dim lsApellido As String, lsNombreLargo As String
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
Dim lmDetalle(20) As String
Dim pnInteresDif As Double, pnMontoRec As Double
Dim ofun As New COMFunciones.FCOMImpresion
Dim lnTicTEA As Double, lnTasInt As Double '*** PEAC 20101220
'lnTicTEA = Round((((1 + lnTasInt / 100) ^ 12 - 1) * 100), 2)

Dim cTempoTexto75 As String * 75

Dim lnSum14k As Double
Dim lnSum16k As Double
Dim lnSum18k As Double
Dim lnSum21k As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

pnInteresDif = pnInteres

If pbBuscaData = True Then

    Dim lrCredPig As New ADODB.Recordset
    Dim lrCredPigJoyas As New ADODB.Recordset
    Dim lrCredPigPersonas As New ADODB.Recordset
    Dim lrCredPigCostos As New ADODB.Recordset
    Dim lrCredPigJoyasDet As New ADODB.Recordset
    Dim lrPigTasas As New ADODB.Recordset
    
    Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato
    
    Set lrPigTasas = loMuestraContrato.dObtieneDatosCreditoPignoraticioTasas(psCodCta)
   
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
        psmensaje = " No se encuentra el Credito Pignoraticio " & psCodCta
        Exit Function
    Else
    
        '***Modificado por ELRO el 20120110, según Acta N° 002-2012/TI-D
        'lnTasInt = lrPigTasas("nTasaInteres")
        lnTasInt = lrCredPig!nTasaInteres
        '***Fin Modificado por ELRO*************************************
        
        pnOroBruto = lrCredPig!nOroBruto
        pnOroNeto = lrCredPig!nOroNeto
        pnPiezas = lrCredPig!nPiezas
        pnValorTasacion = lrCredPig!nTasacion
        pnMontoPrestado = lrCredPig!nMontoCol
        pnInteresDif = IIf(IsNull(lrCredPig!nIntDif), pnInteres, lrCredPig!nIntDif)
        
        '*** PEAC 20071207 - solo muestra el interes no calcula ********
        'pnMontoRec = pnMontoPrestado - pnInteresDif
        pnMontoRec = pnMontoPrestado
        '****************************************************************
        
        
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
        
        'pnImpuesto = lrCredPigCostos!nImpuesto
        pnImpuesto = pnITF
        
        lrCredPigCostos.Close
        Set lrCredPigCostos = Nothing
        Dim item As Integer, itemTMP As Integer
        item = 0
        Do While Not lrCredPigJoyasDet.EOF
            lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 31) & " " _
                        & lrCredPigJoyasDet!cKilataje & "K " & Space(16) & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
            If Trim(lrCredPigJoyasDet!cKilataje) = 14 Then
                lnSum14k = lnSum14k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 16 Then
                lnSum16k = lnSum16k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 18 Then
                lnSum18k = lnSum18k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 21 Then
                lnSum21k = lnSum21k + lrCredPigJoyasDet!npesoneto
            End If
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
'''lsSM = "S/. " 'marg ers044-2016
lsSM = gcPEN.gcPEN_SIMBOLO & " " 'marg ers044-2016
lsCadImp = ""

lsDescLote = lsDescLote & Chr$(13)
        
prPersonas.MoveFirst

Dim sSql As String, RSTEMP As ADODB.Recordset, cn As COMConecta.DCOMConecta

 Set RSTEMP = New ADODB.Recordset
 Set cn = New COMConecta.DCOMConecta
  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre,cPersNombre as cPersNombreLargo "
  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
  If RSTEMP.State = 1 Then RSTEMP.Close
  cn.AbreConexion
  Set RSTEMP = cn.CargaRecordSet(sSql)
  If RSTEMP.State = 1 Then
    If RSTEMP.RecordCount > 0 Then
       lsApellido = RSTEMP!cPersApellido
       lsNombre = RSTEMP!cPersNombre
       lsNombreLargo = RSTEMP!cPersNombreLargo
    End If
    If RSTEMP.State = 1 Then RSTEMP.Close
  End If
  cn.CierraConexion
  Set cn = Nothing

    lsCodPers = prPersonas!cPersCod
    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
    lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)


    lnTicTEA = Round((((1 + lnTasInt / 100) ^ 12 - 1) * 100), 2)
'   lnTimTEA = Round((((1 + lnTasMor / 100) ^ 12 - 1) * 100), 2)

    lsCadImp = ""
'    lsCadImp = Chr$(27) & Chr$(64)
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) 'Condensada
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(103)
'    lsCadImp = lsCadImp & "   " & oFunI.gPrnSaltoLinea 'oFunI.gPrnTamLetra12CPI
'    lsCadImp = lsCadImp & "   " & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)
    
    '**Configuración de página**********************************************************
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(66)   'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    'lsCadImp = lsCadImp & Chr$(27) + Chr$(18)  ' cancela condensada
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    '***********************************************************************************
    
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 8)
    lsCadImp = lsCadImp & Space(20) & lsCodPers & Space(21) & Format(psFecContrato, "dd/MM/yyyy") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(54) & Format(psFecVencimiento, "dd/MM/yyyy") & Space(14)
   
    lsCadImp = lsCadImp & oFunI.gPrnBoldON & psCodCta & oFunI.gPrnBoldOFF & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNombreLargo) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(48) & ofun.ImpreCarEsp(lsNombre) & oFunI.gPrnSaltoLinea
    'cTempoTexto75 = ""
    'cTempoTexto75 = ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario)
    'lsCadImp = lsCadImp & Space(48) & cTempoTexto75 & oFunI.gPrnSaltoLinea  '& Space(24) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreFormat(ofun.ImpreCarEsp(lsDireccion), 50, 0) & Space(7) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnPiezas, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoPrestado, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
    '*** PEAC 20101220
    'lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroBruto, 12, 2, True) & Space(34) & ofun.ImpreFormat(pnTasaInteresMensual, 12, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroBruto, 12, 2, True) & Space(34) & ofun.ImpreFormat(lnTicTEA, 12, 2) & oFunI.gPrnSaltoLinea
    'FIN PEAC
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroNeto, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnInteresDif, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(26) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoRec - pnImpuesto, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoRec, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    'Titulo y Monto Prestamo
    cTempoTexto75 = ""
    'cTempoTexto75 = "Item" & Space(3) & "Descripción" & Space(22) & "P. Neto"    '& Space(21) &oFun.ImpreFormat(pnInteresDif, 9, 2, True)
    
    'Primer Item y el Prestamo
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(0), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnMontoPrestado, 18, 2, True) & oFunI.gPrnSaltoLinea
      
    'Segundo Item y interes
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(1), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnInteresDif, 18, 2, True) & oFunI.gPrnSaltoLinea
    
    'tercer Item y Costo de Tasacion
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(2), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoTasacion, 18, 2, True) & oFunI.gPrnSaltoLinea  '4
    
    'Cuarto Item y Costo de Custodia
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(3), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoCustodia, 18, 2, True) & oFunI.gPrnSaltoLinea  '5
    
    'Quinto Item y el Impuesto
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(4), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea  '6
    
    'Item 6
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(5), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    
    'Item 7
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(6), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
        
    'Item 8
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(7), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    
    'Item 9
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(8), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & Space(22) & oFunI.gPrnSaltoLinea
        
    'Item 10 y Neto a Recibir
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(9), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(18) & ofun.ImpreFormat(pnMontoRec - pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea
     
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(10), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(11), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(12), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(13), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(14), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(15), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(16), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(17), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(18), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(19), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea

'    lsCadImp = lsCadImp & Space(20) & "SON: " & lnSum14k & " GR.14K. " & lnSum18k & " GR.18K. "
'    lsCadImp = lsCadImp & Space(1) & lnSum21k & " GR.21K. " & lnSum16k & " GR.16K. " & oFunI.gPrnSaltoLinea

'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(16), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(17), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(18), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(19), 55)
'    lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(60) & "T.E.A Adelantada 42.766 % " & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & Space(70) & "Nro Piezas: " & ofun.ImpreFormat(pnPiezas, 18, 2, True)
'
'    'Cambio de Tamaño de letra a 12
'    lsCadImp = lsCadImp & oFunI.gPrnTamLetra15CPI & oFunI.gPrnBoldON
'    lsCadImp = lsCadImp & Space(10) & psCodCta & oFunI.gPrnSaltoLinea
'     'Cambio de Tamaño de letra a 10
'    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF & oFunI.gPrnTamLetra10CPI
'
'    lsCadImp = lsCadImp & Space(70) & "Oro Bruto : " & ofun.ImpreFormat(pnOroBruto, 18, 2, True) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(70) & "Oro Neto  : " & ofun.ImpreFormat(pnoroneto, 18, 2, True) & Space(30) & pnPlazo & oFunI.gPrnSaltoLinea
'
'
'    cTempoTexto75 = ""
'    If pnNroDuplicado > 0 Then
'        cTempoTexto75 = "Duplicado Nro : " & ofun.ImpreFormat(pnNroDuplicado, 5, 0)
'    End If
    
    'lsCadImp = lsCadImp & Space(20) & cTempoTexto75 & Space(1) & psCodUser & Space(25) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
    
    'lsCadImp = lsCadImp & Chr$(27) '+ Chr$(18) ' cancela condensada

    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 7)
    lsCadImp = lsCadImp & Space(71) & oFunI.gPrnBoldON & psCodCta & oFunI.gPrnBoldOFF & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(15) & "Venc." & Format(psFecVencimiento, "dd/MM/yyyy") & Space(10) & psCodUser & Space(33) & pnPlazo
    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
    nPrintContratoPignoraticioDet = lsCadImp
    Set oFunI = Nothing
End Function

'Procedimiento de impresión del recibo de renovación
Public Function nPrintReciboRenovacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnSaldoCapAnt As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnInteresMor As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, ByVal psVarPersCMAC As String, _
        ByVal psGastoCorrespondencia As Currency, Optional ByVal psmensaje As String = "", Optional ByVal pnITF As Double = 0#, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pnInteresVen As Currency = 0, _
        Optional ByVal pbImpTMU As Boolean = False, Optional ByVal gbAntNueRevol As Boolean = True, _
        Optional ByVal fnPorcentajePagoCapital As Double = 0, _
        Optional ByVal pnNumDiasFeriado As Integer = 0, _
        Optional ByVal pnCampana As Integer = -1, Optional ByVal pnAplicaCampa As Integer = -1) As String
        '*** Jony M. - se gregó los parámetros gbAntNueRevol y fnPorcentajePagoCapital
        '*** PEAC 20190712 - SE AGREGÓ pnNumDiasFeriado
        'JOEP20210921 campana prendario pnCampana,pnAplicaCampa 0:No aplica,1:aplica
        
        
    Dim lsCadImp As String
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim loConstSistema As New COMDConstSistema.NCOMConstSistema
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
    Dim vEspacio As Integer
    
    Dim nPorcenCapital As Double
    
    'peac 20070820
    
'    'Para verificar la fecha de Vencimiento -- AVMM --11-06-2007************
'    Dim oCon As COMConecta.DCOMConecta
'    Dim sSql As String
'    Dim Rs As ADODB.Recordset
'    Dim ldVarFecVencimiento As Date
'    Dim ldFechaAct As Date
'    Dim lnNewPlazo As Integer
'    Set oCon = New COMConecta.DCOMConecta
'        If oCon.AbreConexion Then
'            sSql = "select dVenc,nPlazo from Colocaciones where cCtaCod='" & psCodCta & "' "
'            Set Rs = New ADODB.Recordset
'            Set Rs = oCon.CargaRecordSet(sSql)
'                If Not (Rs.EOF And Rs.BOF) Then
'                    ldVarFecVencimiento = Rs("dVenc")
'                    ldFechaAct = Format(psFecTransac, "dd/mm/yyyy")
'                    lnNewPlazo = Rs("nPlazo")
'                End If
'            Set Rs = Nothing
'
'        End If
'    Set oCon = Nothing
'    '***********************************************************************
'    psFecProxPago = Format$(ldVarFecVencimiento, "dd/mm/yyyy")
    
    ' end peac ***********************
    
'    If Format(psFecProxPago, "dd/mm/yyyy") <> Format(ldVarFecVencimiento, "dd/mm/yyyy") Then
'        psFecProxPago = Format$(ldFechaAct + lnNewPlazo, "dd/mm/yyyy")
'    End If

    If pbImpTMU = False Then
        vEspacio = 6
        lsCadImp = "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Renovacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(1) & Space(22 + vEspacio) _
                            & "Pignoraticio-Renovacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio + 8) _
                            & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(9 + vEspacio) _
                                 & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & "Contrato  : " & lsCodAntiguo & Space(10) & Space(9 + vEspacio) _
                                & "Contrato  : " & lsCodAntiguo & Space(5) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(9 + vEspacio) _
                                & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(4) & oFunI.gPrnSaltoLinea
        End If
        'oFun.oFun.ImpreFormat(psNombCliente, 28)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(48 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(34 - Len(Left(Trim(psNombCliente), 30))) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio + 8) _
                            & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & " DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
        
        'Comentado Por MAVM 20140320 ***
        'lsCadImp = lsCadImp & "Saldo          :" & ofun.ImpreFormat(pnSaldoCapAnt, 21, , True) & Space(vEspacio + 10) _
        '                    & "Saldo          :" & ofun.ImpreFormat(pnSaldoCapAnt, 21, , True) & oFunI.gPrnSaltoLinea
        '***
        
        lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 10) _
                            & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Mora           :" & ofun.ImpreFormat(pnInteresMor, 21, , True) & Space(vEspacio + 10) _
                            & "Mora           :" & ofun.ImpreFormat(pnInteresMor, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 10) _
                            & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Int.Vencido    :" & ofun.ImpreFormat(pnInteresVen, 21, , True) & Space(vEspacio + 10) _
                            & "Int.Vencido    :" & ofun.ImpreFormat(pnInteresVen, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia + pnImpuesto, 21, , True) & Space(vEspacio + 10) _
                            & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia + pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        'Comentado Por MAVM 20140320 ***
        'lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 10) _
        '                    & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
        '***
                            
        lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & Space(vEspacio + 10) _
                            & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 10) _
                            & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 10) _
                            & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio + 10) _
                            & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
        'RECO20140606 ERS044-2014****************************************************************************************
        
        If gbAntNueRevol = True Then ' true =SI ES ANTIGUO
            nPorcenCapital = 0
        Else
            nPorcenCapital = Round(pnNewSaldoCap * fnPorcentajePagoCapital, 2) '*** 20171109 - JHONY M.
        End If
        
        'pnInteres = (pnNewSaldoCap * (pnTasaInteresMensual / 100)) 'RECO20140825
        'pnInteres = Round(pnNewSaldoCap * (pnTasaInteresMensual / 100), 2) + nPorcenCapital '*** 20171109 - JHONY M.
        pnInteres = Round((((1 + pnTasaInteresMensual / 100) ^ ((pnPlazo + pnNumDiasFeriado) / 30)) - 1) * pnNewSaldoCap, 2) + nPorcenCapital '*** PEAC 20190711
        
        'lsCadImp = lsCadImp & "NUEVO INTERES  :" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 10) _
                            & "NUEVO INTERES  :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
        '*** 20171109 - JHONY M.
        lsCadImp = lsCadImp & "NVO. MONTO PAGO:" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 10) _
                            & "NVO. MONTO PAGO:" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
                            
        'RECO FIN********************************************************************************************************
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Renovación Nro : ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ofun.ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio + 10) _
                            & "" & ofun.ImpreCarEsp("Renovación Nro : ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ofun.ImpreFormat(pnPlazo, 4, 0) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & Space(8 + vEspacio) _
                                 & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        'comentado por peac 20070820
        'lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Space(vEspacio + 8) _
                            & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                            
        'MAVM 20140304*** Comentado x JUEZ 20140515
        'If (Mid(psCodCta, 4, 2)) = "01" Or (Mid(psCodCta, 4, 2) = "04") Or (Mid(psCodCta, 4, 2) = "09") Or (Mid(psCodCta, 4, 2) = "31") Then
        '    lsCadImp = lsCadImp & "Estimado Cliente de credito prendario, a partir" & Space(vEspacio) & "Estimado Cliente de credito prendario, a partir" & oFunI.gPrnSaltoLinea _
        '                        & "del 15/05/14 el costo por aviso notarial de venta" & Space(4) & "del 15/05/14 el costo por aviso notarial de venta" & oFunI.gPrnSaltoLinea _
        '                        & "o adjudicacion de su garantia sera de S/. 17.00" & Space(vEspacio) & "o adjudicacion de su garantia sera de S/. 17.00" & oFunI.gPrnSaltoLinea
        'End If
        '***
        
    'JOEP20210921 campana prendario
        If (pnCampana = 1) And pnNroRenovacion = 1 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Por tu pago, podrás acceder al beneficio") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Por tu pago, podrás acceder al beneficio") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("de descuento en tu próxima renovación.") & Space(12 + vEspacio) & ofun.ImpreCarEsp("de descuento en tu próxima renovación.") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Sujeto a los términos y condiciones de") & Space(12 + vEspacio) & ofun.ImpreCarEsp("Sujeto a los términos y condiciones de") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("la campaña.") & Space(39 + vEspacio) & ofun.ImpreCarEsp("la campaña.")
        ElseIf (pnCampana = 1) And pnNroRenovacion = 2 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("¡FELICITACIONES!") & Space(34 + vEspacio) & ofun.ImpreCarEsp("¡FELICITACIONES!") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 20%") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 20%") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("en el pago de tu renovación.") & Space(22 + vEspacio) & ofun.ImpreCarEsp("en el pago de tu renovación.")
        ElseIf (pnCampana = 1) And pnNroRenovacion = 3 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("¡FELICITACIONES!") & Space(34 + vEspacio) & ofun.ImpreCarEsp("¡FELICITACIONES!") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 30%") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 30%") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("en el pago de tu renovación.") & Space(22 + vEspacio) & ofun.ImpreCarEsp("en el pago de tu renovación.")
        End If
    'JOEP20210921 campana prendario
        
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 8) _
                            & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
       
        
    Else
    
    vEspacio = 6
        'lsCadImp = "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "PPignoraticio-Renovacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(1) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & "CContrato  : " & lsCodAntiguo & Space(10) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "CContrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
        End If
        'oFun.oFun.ImpreFormat(psNombCliente, 28)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(48 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Fec.Prestamo:" & Format(psFecPrestamo, "dd/mm/yyyy") & " DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Saldo          :" & ofun.ImpreFormat(pnSaldoCapAnt, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Mora           :" & ofun.ImpreFormat(pnInteresMor, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Int.Vencido    :" & ofun.ImpreFormat(pnInteresVen, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia + pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "II.T.F.        :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
        'RECO20140606 ERS044-2014****************************************************************************************
        
        If gbAntNueRevol = True Then ' true =SI ES ANTIGUO
            nPorcenCapital = 0
        Else
            nPorcenCapital = Round(pnNewSaldoCap * fnPorcentajePagoCapital, 2) '*** 20171109 - JHONY M.
        End If
        
        'pnInteres = (pnNewSaldoCap * (pnTasaInteresMensual / 100)) 'RECO20140825
        'pnInteres = Round(pnNewSaldoCap * (pnTasaInteresMensual / 100), 2) + nPorcenCapital '*** 20171109 - JHONY M.
         pnInteres = Round((((1 + pnTasaInteresMensual / 100) ^ ((pnPlazo + pnNumDiasFeriado) / 30)) - 1) * pnNewSaldoCap, 2) + nPorcenCapital '*** PEAC 20190711

        
        'lsCadImp = lsCadImp & "NUEVO INTERES    :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
        
        '*** 20171109 - JHONY M.
        lsCadImp = lsCadImp & "NVO. MONTO PAGO:" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
        
        'RECO FIN********************************************************************************************************
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Renovación Nro : ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " & ofun.ImpreFormat(pnPlazo, 4, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        'lsCadImp = lsCadImp & "Tasa Efectiva Anual:" & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                          
        'MAVM 20140304*** Comentado x JUEZ 20140515
        'If (Mid(psCodCta, 4, 2)) = "01" Or (Mid(psCodCta, 4, 2) = "04") Or (Mid(psCodCta, 4, 2) = "09") Or (Mid(psCodCta, 4, 2) = "31") Then
        '    lsCadImp = lsCadImp & "Estimado Cliente de credito prendario, a" & oFunI.gPrnSaltoLinea _
        '                        & "partir del 15/05/14 el costo por aviso " & oFunI.gPrnSaltoLinea _
        '                        & "notarial de venta o adjudicacion de su " & oFunI.gPrnSaltoLinea _
        '                        & "garantia sera de S/. 17.00" & oFunI.gPrnSaltoLinea
        'End If
        '***
    'JOEP20210921 campana prendario
        If (pnCampana = 1) And pnNroRenovacion = 1 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Por tu pago, podrás acceder al beneficio") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Por tu pago, podrás acceder al beneficio") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("de descuento en tu próxima renovación.") & Space(12 + vEspacio) & ofun.ImpreCarEsp("de descuento en tu próxima renovación.") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Sujeto a los términos y condiciones de") & Space(12 + vEspacio) & ofun.ImpreCarEsp("Sujeto a los términos y condiciones de") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("la campaña.") & Space(39 + vEspacio) & ofun.ImpreCarEsp("la campaña.")
        ElseIf (pnCampana = 1) And pnNroRenovacion = 2 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("¡FELICITACIONES!") & Space(34 + vEspacio) & ofun.ImpreCarEsp("¡FELICITACIONES!") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 20%") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 20%") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("en el pago de tu renovación.") & Space(22 + vEspacio) & ofun.ImpreCarEsp("en el pago de tu renovación.")
        ElseIf (pnCampana = 1) And pnNroRenovacion = 3 And pnAplicaCampa = 1 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("¡FELICITACIONES!") & Space(34 + vEspacio) & ofun.ImpreCarEsp("¡FELICITACIONES!") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 30%") & Space(10 + vEspacio) & ofun.ImpreCarEsp("Accediste a la tasa de descuento del 30%") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("en el pago de tu renovación.") & Space(22 + vEspacio) & ofun.ImpreCarEsp("en el pago de tu renovación.")
        End If
    'JOEP20210921 campana prendario
                          
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        
    End If
    
    nPrintReciboRenovacion = lsCadImp
    Set oFunI = Nothing
End Function

'Procedimiento de impresión del recibo de cancelacion del contrato
Public Function nPrintReciboCancelacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnMontoPrestamo As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteresAdel As Currency, ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal psCodUser As String, ByVal pnCostoNotificacion As Currency, _
        ByVal psVarPersCMAC As String, Optional ByVal psmensaje As String = "", Optional psGastoCorrespondencia As Currency, _
        Optional ByVal pnITF As Double = 0#, Optional ByVal psImpresora As Impresoras = gEPSON, _
        Optional ByVal pnIntMora As Currency = 0, Optional ByVal pbImpTMU As Boolean = False) As String
    
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora
    Dim vEspacio As Integer
    
    If pbImpTMU = False Then
        vEspacio = 6
        lsCadImp = "" & oFunI.gPrnSaltoLinea  '& oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Pignoraticio-Cancelacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(10 + vEspacio) _
                            & "Pignoraticio-Cancelacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        If psVarPersCMAC <> "" Then
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio) _
                                & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        End If
        'oFun.oFun.ImpreFormat(psNombCliente, 28)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(1 + vEspacio) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio) _
                            & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Prestamo       :" & ofun.ImpreFormat(pnMontoPrestamo, 21, , True) & Space(vEspacio + 2) _
                            & "Prestamo       :" & ofun.ImpreFormat(pnMontoPrestamo, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 2) _
                            & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteresAdel, 21, , True) & Space(vEspacio + 2) _
                            & "Interes        :" & ofun.ImpreFormat(pnInteresAdel, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Int. Vencido   :" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 2) _
                            & "Int. Vencido   :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Int. Moratorio :" & ofun.ImpreFormat(pnIntMora, 21, , True) & Space(vEspacio + 2) _
                            & "Int. Moratorio :" & ofun.ImpreFormat(pnIntMora, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio + 2) _
                            & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
                                
'       lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & oFunI.gPrnSaltoLinea
                                
        'peac 20071220
        lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(pnCostoNotificacion, 21, , True) & Space(vEspacio + 2) _
                            & "Costo Notif.   :" & ofun.ImpreFormat(pnCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 2) _
                            & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "SALDO DE CAPITAL :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio) _
                            & "SALDO DE CAPITAL :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & Space(vEspacio) _
                            & "" & ofun.ImpreCarEsp("Cancelación Pignoraticio ") & Space(13) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & " " & oFun.ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(10 + vEspacio) & oFun.ImpreCarEsp("Próxima fecha de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
        'comentado por peac 20070820
        'lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(13 + vEspacio) _
                            & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 38) & Space(vEspacio) _
                            & "" & ofun.ImpreFormat(psmensaje, 38) & oFunI.gPrnSaltoLinea
    Else
        'lsCadImp = "" & oFunI.gPrnSaltoLinea  '& oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "PPignoraticio-Cancelacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & LTrim(ofun.ImpreFormat(psNomAgencia, 24, 0)) & Space(1) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        If psVarPersCMAC <> "" Then
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        End If
        'oFun.oFun.ImpreFormat(psNombCliente, 28)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "CContrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Fec.Prestamo:" & Format(psFecPrestamo, "dd/mm/yyyy") & " DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & "Prestamo       :" & ofun.ImpreFormat(pnMontoPrestamo, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteresAdel, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Int. Vencido   :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Int. Moratorio :" & ofun.ImpreFormat(pnIntMora, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
                                
'       lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psGastoCorrespondencia, 21, , True) & oFunI.gPrnSaltoLinea
                                
        lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(pnCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea
                                
                                
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & "II.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "SALDO DE CAPITAL :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Cancelación Pignoraticio ") & Space(6) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        'comentado por peac 20070820
        'lsCadImp = lsCadImp & "Tasa Efectiva Anual:" & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 38) & Space(vEspacio) _
                            & "" & ofun.ImpreFormat(psmensaje, 38) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    End If
    nPrintReciboCancelacion = lsCadImp
    Set oFunI = Nothing
End Function


Public Function nPrintReciboDesembolso(ByVal psFecVen As String, ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal pnMontoEntregar As Currency, ByVal pnInteresCompensat As Currency, ByVal psNomAgencia As String, _
        ByVal psCodUser As String, ByVal pnITF As Double, Optional ByVal psImpresora As Impresoras = gEPSON, _
        Optional ByVal pbDesembCC As Boolean = False, Optional ByVal psCtaAho As String = "", Optional ByVal psCtaAhoNew As String = "", Optional ByVal pbImpTMU As Boolean = False) As String
        
' PEAC 20090513 - Se agrego el parametro pbDesembCC que indica si es con o sin abono en cuenta
        
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim psNombCliente As String
Dim Co As COMConecta.DCOMConecta
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim lsTit As String, sText As String, sIni As String, lsTextPig As String, sText1 As String ''*** PEAC 20090513
Dim sText2 As String, sText3 As String
Dim lnNetoRecibir As Currency, lsSpace As Integer, lsSpace1 As Integer ''** PEAC 20090513

Set Co = New COMConecta.DCOMConecta
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

sql = " select dbo.pstaNombre(cPersNombre,0) Nombre from Persona PE " & _
      " Inner Join ProductoPersona PP on PE.cPersCod = PP.cPersCod " & _
      " Where nPrdPersRelac =" & ColocRelacPers.gColRelPersTitular & " And cCtaCod ='" & psCtaCod & "'"

Co.AbreConexion
Set Rs = Co.CargaRecordSet(sql)
Co.CierraConexion
psNombCliente = Rs!NOMBRE
Set Rs = Nothing
Set Co = Nothing
    Dim vEspacio As Integer
    vEspacio = 6
    '--oFun.oFun.ImpreFormat(psNombCliente, 28)
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita

        ''*** PEAC 20090516
        If pbDesembCC Then
            lsTextPig = "Pignoraticio-Desembolso con Abono a Cuenta"
            'lnNetoRecibir = pnSaldoCap
            lnNetoRecibir = pnMontoEntregar ' RECO20140912
            lsSpace = 0
            lsSpace1 = 4
            
            If Len(psCtaAhoNew) > 0 Then
                lsTit = "----    APERTURA AHORROS    ----"
            Else
                lsTit = "----    DEPOSITO AHORROS    ----"
            End If
            
            
        Else
            lsTextPig = "Pignoraticio-Desembolso"
            lnNetoRecibir = pnMontoEntregar
            lsSpace = 14
            lsSpace1 = 9
        End If
        
        '******************************add pti1 11/03/2019 ACTA 030-2019*******************************************
            If pbImpTMU = False Then
            '*NO TMU********************************************************
                lsCadImp = lsCadImp & lsTextPig & Space(lsSpace1) & Space(lsSpace + vEspacio) _
                                    & lsTextPig & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
                lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & psFechaHora & Space(7) _
                                    & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(1) & psFechaHora & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
                lsCadImp = lsCadImp & "Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & Space(5 + vEspacio) _
                                    & "Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(10) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) _
                                    & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Fec.Prestamo: " & psFechaHora & "             " & Space(vEspacio) _
                                    & "Fec.Prestamo: " & psFechaHora & "             " & oFunI.gPrnSaltoLinea
                
                lsCadImp = lsCadImp & "Prestamo          :" & ofun.ImpreFormat(pnSaldoCap, 18, , True) & Space(vEspacio + 6) _
                                    & "Prestamo          :" & ofun.ImpreFormat(pnSaldoCap, 18, , True) & oFunI.gPrnSaltoLinea
                                    
                lsCadImp = lsCadImp & "Interes Comp.Vcmto:" & ofun.ImpreFormat(pnInteresCompensat, 18, , True) & Space(vEspacio + 6) _
                                    & "Interes Comp.Vcmto:" & ofun.ImpreFormat(pnInteresCompensat, 18, , True) & oFunI.gPrnSaltoLinea
    
                If Not pbDesembCC Then
                    lsCadImp = lsCadImp & "I.T.F.            :" & ofun.ImpreFormat(pnITF, 18, , True) & Space(vEspacio + 6) _
                                        & "I.T.F.            :" & ofun.ImpreFormat(pnITF, 18, , True) & oFunI.gPrnSaltoLinea
                End If
            
                lsCadImp = lsCadImp & Chr$(27) & Chr$(69)
                
                lsCadImp = lsCadImp & "NETO A RECIBIR    :" & ofun.ImpreFormat(lnNetoRecibir, 18, , True) & Space(vEspacio + 6) _
                                    & "NETO A RECIBIR    :" & ofun.ImpreFormat(lnNetoRecibir, 18, , True) & oFunI.gPrnSaltoLinea 'peac 20070820 & oFunI.gPrnSaltoLinea
                       
                lsCadImp = lsCadImp & "Próxima fecha de pago: " & Format(psFecVen, "dd/mm/yyyy") & Space(14 + vEspacio) _
                                    & "Próxima fecha de pago: " & Format(psFecVen, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
                
            
                If (Mid(psCtaCod, 4, 2)) <> "01" And (Mid(psCtaCod, 4, 2) <> "04") And (Mid(psCtaCod, 4, 2) <> "09") And (Mid(psCtaCod, 4, 2) <> "31") Then 'MAVM 20140331***
                    lsCadImp = lsCadImp & "" & Space(37) & Format(psCodUser, "@@@@") & Space(vEspacio + 6) _
                                        & "" & Space(37) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
                Else
                        lsCadImp = lsCadImp & "" & Space(37) & Format(psCodUser, "@@@@") & Space(vEspacio + 6) _
                                            & "" & Space(37) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)
                End If
               
                If pbDesembCC Then
        
                        sIni = 0
                        
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)  'Activa Negrita
                        lsCadImp = lsCadImp & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Space(11) & Space(14 + vEspacio) & _
                                                            "CMAC-MAYNAS - AHORROS" & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
                        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & psFechaHora & Space(7) _
                                            & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(1) & psFechaHora & oFunI.gPrnSaltoLinea
                
                        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) _
                                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & Space(sIni) & "Cuenta:" & psCtaAho & Space(8) & Space(13 + vEspacio) & _
                                                            "Cuenta:" & psCtaAho & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)  'Activa Negrita
                        lsCadImp = lsCadImp & Space(5) & lsTit & Space(9) & Space(0 + vEspacio) & _
                                              Space(5) & lsTit & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)  'Desactiva Negrita
                                        
                        sText = "Deposito OP"
                        sText1 = "ITF Efectivo"
                        sText2 = "Saldo Disponible"
                        sText3 = "Saldo Contable"
                        
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & _
                                              Space(-1 + 14) & (Trim(Mid(sText, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                                              
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText1, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText1, 1, 28))) - Len(pnITF)) & Format(pnITF, "0.00") & _
                                              Space(-1 + 14) & (Trim(Mid(sText1, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText1, 1, 28))) - Len(pnITF)) & Format(pnITF, "0.00") & oFunI.gPrnSaltoLinea
        
                        If Len(psCtaAhoNew) > 0 Then
                            lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText2, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText2, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & _
                                                  Space(-1 + 14) & (Trim(Mid(sText2, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText2, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea
            
                            lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText3, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText3, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & _
                                                  Space(-1 + 14) & (Trim(Mid(sText3, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText3, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea
                        End If
                        
                        lsCadImp = lsCadImp & "---------------------------------------" & Space(7 + vEspacio) _
                                            & "---------------------------------------" & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(Format(psCodUser, "@@@@"), 1, 28))) & Space(37 + 6 - Len(Trim(Mid(Format(psCodUser, "@@@@"), 1, 28)))) & Space(-1 + 10) & _
                                                            (Trim(Mid(Format(psCodUser, "@@@@"), 1, 28))) & Space(37 + 6 - Len(Trim(Mid(Format(psCodUser, "@@@@"), 1, 28)))) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
                End If
            Else
        '*SI ES TMU ************************************************** MODIFICADO POR PTI1 08-04-2019
                lsCadImp = lsCadImp & " " & lsTextPig & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
                lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
                lsCadImp = lsCadImp & " " & "Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Fec.Prestamo: " & psFechaHora & "             " & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Prestamo          :" & ofun.ImpreFormat(pnSaldoCap, 18, , True) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Interes Comp.Vcmto:" & ofun.ImpreFormat(pnInteresCompensat, 18, , True) & oFunI.gPrnSaltoLinea
    
                If Not pbDesembCC Then
                    lsCadImp = lsCadImp & "I.T.F.            :" & ofun.ImpreFormat(pnITF, 18, , True) & oFunI.gPrnSaltoLinea
                End If
                
                lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
                lsCadImp = lsCadImp & " " & "NETO A RECIBIR    :" & ofun.ImpreFormat(lnNetoRecibir, 18, , True) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & " " & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Próxima fecha de pago: " & Format(psFecVen, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & " " & Space(35) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & " " & oFunI.gPrnSaltoLinea
                'para que corte la TMU
                lsCadImp = lsCadImp & "" & Space(30) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        
                If pbDesembCC Then
        
                        sIni = 0
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)  'Activa Negrita
                        lsCadImp = lsCadImp & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Space(11) & Space(14 + vEspacio) & _
                                                            "CMAC-MAYNAS - AHORROS" & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
                        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & psFechaHora & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Space(sIni) & "Cuenta:" & psCtaAho & Space(8) & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)  'Activa Negrita
                        lsCadImp = lsCadImp & Space(5) & lsTit & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)  'Desactiva Negrita
                                        
                        sText = "Deposito OP"
                        sText1 = "ITF Efectivo"
                        sText2 = "Saldo Disponible"
                        sText3 = "Saldo Contable"
                        
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText1, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText1, 1, 28))) - Len(pnITF)) & Format(pnITF, "0.00") & oFunI.gPrnSaltoLinea
        
                        If Len(psCtaAhoNew) > 0 Then
                            lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText2, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText2, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(sText3, 1, 28))) & Space(35 + 6 - Len(Trim(Mid(sText3, 1, 28))) - Len(lnNetoRecibir)) & Format(lnNetoRecibir, "0.00") & oFunI.gPrnSaltoLinea
                        End If
                        
                        lsCadImp = lsCadImp & "---------------------------------------" & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & Space(sIni) & (Trim(Mid(Format(psCodUser, "@@@@"), 1, 28))) & Space(37 + 6 - Len(Trim(Mid(Format(psCodUser, "@@@@"), 1, 28)))) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
                End If
            End If
            
        nPrintReciboDesembolso = lsCadImp
        Set oFunI = Nothing
        'END PTI1 ACTA 030-2019
            
'***************************************************ELIMINADO POR PTI1 ACTA 030-2019 ********************************************************
'        lsCadImp = lsCadImp & lsTextPig & Space(lsSpace1) & Space(lsSpace + vEspacio) .../*
'        */....
'    Set oFunI = Nothing
'***************************************************FIN  ELIMINADO Y COMENTADO PTI1****************************************
End Function

Public Function nPrintReciboDuplicadoContrato(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnCostoDuplicado As Currency, _
        ByVal pnNroDuplicado As Integer, pnTasaInteresMensual As Double, ByVal psCodUser As String, _
        ByVal psmensaje As String, ByVal pnITF As Double, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Credito Pignoraticio " & Space(11) & Space(9 + vEspacio) _
                            & "Credito Pignoraticio " & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio - 1) _
                            & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio - 1) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & "--------------------------------------" & Space(2 + vEspacio) & " --------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Costo Duplicado  : " & ofun.ImpreFormat(pnCostoDuplicado, 15) & Space(5 + vEspacio) _
                            & "Costo Duplicado  : " & ofun.ImpreFormat(pnCostoDuplicado, 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "I.T.F.           : " & ofun.ImpreFormat(CDbl(pnITF), 15) & Space(5 + vEspacio) _
                            & "I.T.F.           : " & ofun.ImpreFormat(CDbl(pnITF), 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total Duplicado  : " & ofun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(5 + vEspacio) _
                            & "Total Duplicado  : " & ofun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15)
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & "" & ofun.ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Space(4 + vEspacio) _
                            & "" & ofun.ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Duplicado Contrato ") & Space(23) & Space(vEspacio) _
                            & "" & ofun.ImpreCarEsp("Duplicado Contrato ") & Space(20) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "--------------------------------------" & Space(5 + vEspacio) _
                            & "--------------------------------------" & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 1) _
                            & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        nPrintReciboDuplicadoContrato = lsCadImp
        
        Set oFunI = Nothing
End Function

Public Function nPrintReciboCobroCustodia(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnCostoDuplicado As Currency, _
        ByVal pnNroDuplicado As Integer, pnTasaInteresMensual As Double, ByVal psCodUser As String, _
        ByVal psmensaje As String, ByVal pnITF As Double, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "Credito Pignoraticio " & Space(11) & Space(9 + vEspacio) _
                            & "Credito Pignoraticio " & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio - 1) _
                            & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio - 1) _
                            & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) _
                            & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & "--------------------------------------" & Space(2 + vEspacio) & " --------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Cobro de Custodia:" & ofun.ImpreFormat(pnCostoDuplicado, 15) & Space(6 + vEspacio) _
                            & "Cobro de Custodia:" & ofun.ImpreFormat(pnCostoDuplicado, 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "I.T.F.           :" & ofun.ImpreFormat(CDbl(pnITF), 15) & Space(6 + vEspacio) _
                            & "I.T.F.           :" & ofun.ImpreFormat(CDbl(pnITF), 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total A Pagar    :" & ofun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(6 + vEspacio) _
                            & "Total A Pagar    :" & ofun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "--------------------------------------" & Space(5 + vEspacio) _
                            & "--------------------------------------" & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 1) _
                            & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        nPrintReciboCobroCustodia = lsCadImp
    Set oFunI = Nothing
        
End Function
' Impresión del recibo de Amortizacion
Public Function nPrintReciboAmortizacion(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnSaldoCap As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, ByVal psVarPersCMAC As String, Optional ByVal psmensaje As String = "", Optional ByVal pnITF As Double = 0#, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False, Optional ByVal psCostoNotificacion As Currency = 0, Optional ByVal pnIntVencido As Currency = 0, Optional ByVal pnIntMoratorio As Currency = 0, Optional ByVal pnIntPendiente As Currency = 0) As String
        '*** PEAC 20170330 - se agregó tres parámetros
        'pnIntVencido ,pnIntMoratorio ,pnIntPendiente

Dim lsCadImp As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
Dim vEspacio As Integer

If pbImpTMU = False Then
    vEspacio = 6
    lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(15)  '*****activa condensada
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
    lsCadImp = lsCadImp & "Pignoraticio-Amortizacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(13 + vEspacio) _
                        & "Pignoraticio-Amortizacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
                        
    'lsCadImp = lsCadImp & "Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(13 + vEspacio) _
                        & "Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio + 4) _
                        & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
                        
    If psVarPersCMAC <> "" Then
         lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio + 4) _
                             & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
    End If
    'oFun.oFun.ImpreFormat(psNombCliente, 28)
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(5 + vEspacio) _
                        & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) _
                        & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio + 4) _
                        & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
    
    'Comentado Por MAVM 20140320 ***
    'lsCadImp = lsCadImp & "Saldo          :" & ofun.ImpreFormat(pnSaldoCap, 21, , True) & Space(vEspacio + 6) _
    '                    & "Saldo          :" & ofun.ImpreFormat(pnSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
    '***
                        
    lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 6) _
                        & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 6) _
                        & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
    '*** PEAC 20170330
    lsCadImp = lsCadImp & "Int.Moratorio  :" & ofun.ImpreFormat(pnIntMoratorio, 21, , True) & Space(vEspacio + 6) _
                        & "Int.Moratorio  :" & ofun.ImpreFormat(pnIntMoratorio, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Int.Vencido    :" & ofun.ImpreFormat(pnIntVencido, 21, , True) & Space(vEspacio + 6) _
                        & "Int.Vencido    :" & ofun.ImpreFormat(pnIntVencido, 21, , True) & oFunI.gPrnSaltoLinea
    '*** fin peac
    
                        
                        
    lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio + 6) _
                        & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio + 6) _
                        & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                        
    'Comentado Por MAVM 20140320 ***
    'lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & Space(vEspacio + 6) _
    '                    & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
    '***
    lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & Space(vEspacio + 6) _
                        & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    
    'MODIFICACION MPBR
    lsCadImp = lsCadImp & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 6) _
                        & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio + 6) _
                        & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
    '*** PEAC 20170330
    lsCadImp = lsCadImp & "Int. Pendiente :" & ofun.ImpreFormat(pnIntPendiente, 21, , True) & Space(vEspacio + 6) _
                        & "Int. Pendiente :" & ofun.ImpreFormat(pnIntPendiente, 21, , True) & oFunI.gPrnSaltoLinea
    '***FIN PEAC

    'lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 6) _
    '                    & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
    
    
    'ANTES 31082004
    'lsCadImp = lsCadImp & " NETO A PAGAR     :" &oFun.ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio) _
    '                    & " NETO A PAGAR     :" &oFun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
                        
                        
    lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio + 6) _
                        & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        
    'lsCadImp = lsCadImp & " " & oFun.ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " &oFun.ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio) _
    '                    & " " & oFun.ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " &oFun.ImpreFormat(pnPlazo, 4, 0) & oFunI.gPrnSaltoLinea
    
    'PEAC 20070820 SE AGREGO (& Space(5) & Format(psCodUser, "@@@@"))
     lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & Space(vEspacio + 4) _
                              & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    
    'comentado por PEAC 20070820
    'lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & Space(vEspacio + 4) _
                        & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
    'END PEAC
    
    'MAVM 20140304*** Comentado x JUEZ 20140515
    'If (Mid(psCodCta, 4, 2)) = "01" Or (Mid(psCodCta, 4, 2) = "04") Or (Mid(psCodCta, 4, 2) = "09") Or (Mid(psCodCta, 4, 2) = "31") Then
    '    lsCadImp = lsCadImp & "Estimado Cliente de credito prendario, a partir" & Space(vEspacio) & "Estimado Cliente de credito prendario, a partir" & oFunI.gPrnSaltoLinea _
    '                        & "del 15/05/14 el costo por aviso notarial de venta" & Space(4) & "del 15/05/14 el costo por aviso notarial de venta" & oFunI.gPrnSaltoLinea _
    '                        & "o adjudicacion de su garantia sera de S/. 17.00" & Space(vEspacio) & "o adjudicacion de su garantia sera de S/. 17.00" & oFunI.gPrnSaltoLinea
    'End If
    '***
                        
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 4) _
                        & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
                        
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(18)    '**para quitar condensada
Else
    vEspacio = 6
    'lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(15)  '*****activa condensada
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
    lsCadImp = lsCadImp & "Pignoraticio-Amortizacion-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & " Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
                        '& "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
                        
    If psVarPersCMAC <> "" Then
         lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
                             '& " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
    End If
    'oFun.oFun.ImpreFormat(psNombCliente, 28)
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & "CContrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & oFunI.gPrnSaltoLinea
                        '& "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
                        '& "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
                        '& "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & "Saldo          :" & ofun.ImpreFormat(pnSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Saldo          :" & ofun.ImpreFormat(pnSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                        
    '*** PEAC 20170330
    lsCadImp = lsCadImp & "Int.Moratorio  :" & ofun.ImpreFormat(pnIntMoratorio, 21, , True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Int.Vencido    :" & ofun.ImpreFormat(pnIntVencido, 21, , True) & oFunI.gPrnSaltoLinea
    '*** fin peac
    
    
    lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    
    'MODIFICACION MPBR
    lsCadImp = lsCadImp & "II.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
    
    '*** PEAC 20170330
    lsCadImp = lsCadImp & "Int. Pendiente :" & ofun.ImpreFormat(pnIntPendiente, 21, , True) & oFunI.gPrnSaltoLinea
                        
    'lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & Space(vEspacio + 6) _
    '                    & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago + pnITF, 21, , True) & oFunI.gPrnSaltoLinea
    
    
    'ANTES 31082004
    'lsCadImp = lsCadImp & " NETO A PAGAR     :" &oFun.ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio) _
    '                    & " NETO A PAGAR     :" &oFun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
                        
                        
    lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        '& "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        
    'lsCadImp = lsCadImp & " " & oFun.ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " &oFun.ImpreFormat(pnPlazo, 4, 0) & Space(vEspacio) _
    '                    & " " & oFun.ImpreCarEsp("Renovación número: ") & Format(pnNroRenovacion, "@@") & Space(9) & "Plazo : " &oFun.ImpreFormat(pnPlazo, 4, 0) & oFunI.gPrnSaltoLinea
    
     
    'PEAC 20070820 SE AGREGO (& Space(5) & Format(psCodUser, "@@@@"))
    lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    
    'peac 20070820
    'lsCadImp = lsCadImp & "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                        '& "Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(9) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                      
    'MAVM 20140304*** Comentado x JUEZ 20140515
    'If (Mid(psCodCta, 4, 2)) = "01" Or (Mid(psCodCta, 4, 2) = "04") Or (Mid(psCodCta, 4, 2) = "09") Or (Mid(psCodCta, 4, 2) = "31") Then
    '    lsCadImp = lsCadImp & "Estimado Cliente de credito prendario, a" & oFunI.gPrnSaltoLinea _
    '                            & "partir del 15/05/14 el costo por aviso " & oFunI.gPrnSaltoLinea _
    '                            & "notarial de venta o adjudicacion de su " & oFunI.gPrnSaltoLinea _
    '                            & "garantia sera de S/. 17.00" & oFunI.gPrnSaltoLinea
    'End If
    '***
                      
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 4) & oFunI.gPrnSaltoLinea
                        '& "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
                        
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(18)    '**para quitar condensada
End If
nPrintReciboAmortizacion = lsCadImp
Set oFunI = Nothing
End Function

'*** PEAC 20161029
' Impresión del recibo de Pago Parcial - pignoraticio
Public Function nPrintReciboPagoParcial(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecPrestamo As String, _
        ByVal pnDiasAtraso As Integer, ByVal pnSaldoCap As Currency, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnCostoRemate As Currency, ByVal pnMontoPago As Currency, ByVal pnNewSaldoCap As Currency, _
        ByVal pnTasaInteresMensual As Double, ByVal pnNroRenovacion As Integer, ByVal psFecProxPago As String, _
        ByVal psCodUser As String, ByVal pnPlazo As Integer, ByVal psVarPersCMAC As String, Optional ByVal psmensaje As String = "", Optional ByVal pnITF As Double = 0#, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False, Optional ByVal psCostoNotificacion As Currency = 0) As String
Dim lsCadImp As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
Dim vEspacio As Integer

If pbImpTMU = False Then
    vEspacio = 6
    lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
                        
    lsCadImp = lsCadImp & "Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(13 + vEspacio) _
                        & "Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio + 4) _
                        & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
                        
    If psVarPersCMAC <> "" Then
         lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(1 + vEspacio + 4) _
                             & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
    End If
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(5 + vEspacio) _
                        & "Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) _
                        & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(40 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & Space(vEspacio + 4) _
                        & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & Space(vEspacio + 6) _
                        & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & Space(vEspacio + 6) _
                        & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & Space(vEspacio + 6) _
                        & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & Space(vEspacio + 6) _
                        & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & Space(vEspacio + 6) _
                        & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    
    lsCadImp = lsCadImp & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & Space(vEspacio + 6) _
                        & "I.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & Space(vEspacio + 6) _
                        & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & Space(vEspacio + 6) _
                        & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        
     lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & Space(vEspacio + 4) _
                              & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                                
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 4) _
                        & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
                        
Else
    vEspacio = 6
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
    lsCadImp = lsCadImp & " Pignoraticio-Pgo.Anticip.-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
                        
    If psVarPersCMAC <> "" Then
         lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
    End If
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & "CContrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    lsCadImp = lsCadImp & "CLIENTE:" & Left(Trim(psNombCliente), 39) & Space(44 - Len(Left(Trim(psNombCliente), 39))) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec.Prestamo: " & Format(psFecPrestamo, "dd/mm/yyyy") & "   DIAS ATRASO:" & Format(pnDiasAtraso, "@@@") & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & "Saldo          :" & ofun.ImpreFormat(pnSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Capital        :" & ofun.ImpreFormat(pnCapital, 21, , True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Interes        :" & ofun.ImpreFormat(pnInteres, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Impuesto       :" & ofun.ImpreFormat(pnImpuesto, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Custodia :" & ofun.ImpreFormat(pnCostoCustodia, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "Costo Remate   :" & ofun.ImpreFormat(pnCostoRemate, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & "Costo Notif.   :" & ofun.ImpreFormat(psCostoNotificacion, 21, , True) & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    
    lsCadImp = lsCadImp & "II.T.F.         :" & ofun.ImpreFormat(pnITF, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NETO A PAGAR   :" & ofun.ImpreFormat(pnMontoPago, 21, , True) & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & "NUEVO SALDO    :" & ofun.ImpreFormat(pnNewSaldoCap, 21, , True) & oFunI.gPrnSaltoLinea
     
    lsCadImp = lsCadImp & "" & ofun.ImpreCarEsp("Próx. fec. de pago: ") & Format(psFecProxPago, "dd/mm/yyyy") & Space(5) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
                        
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita

    lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 4) & oFunI.gPrnSaltoLinea
                        '& "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
                        
End If
nPrintReciboPagoParcial = lsCadImp
Set oFunI = Nothing
End Function

'Procedimiento de impresión del recibo de Extorno
Public Function nPrintReciboExtorno(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psOperacion As String, _
        ByVal pnMonto As Currency, ByVal pnSaldoCapAnt As Currency, ByVal pnMovExtornado As Long, _
        ByVal psCodUser As String, ByVal psVarPersCMAC As String, _
        Optional ByVal psmensaje As String = "", Optional ByVal psUsuOri As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False, _
        Optional ByVal psOpeCod As String = "0") As String
    'FRHU 20150817: Se agrego el parametro psCodOpe
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim loConstSistema As New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
        
    Dim vEspacio As Integer
    vEspacio = 28
        
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    If pbImpTMU = False Then
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        'FRHU 20150817 OBSERVACION
        'lsCadImp = lsCadImp & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(12) & Space(14 + vEspacio) _
        '                    & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        If psOpeCod = gColRecOpePagTransfFocMacmEfe Or psOpeCod = gColRecOpePagTransfFocMacmVou Or psOpeCod = gColRecOpePagTransfFocMacmChq Then
            lsCadImp = lsCadImp & " Créditos Transf. FOCMAC" & Space(12) & Space(14 + vEspacio) _
                                & " Créditos Transf. FOCMAC" & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(12) & Space(14 + vEspacio) _
                                & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        End If
        'FIN FRHU 20150817
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & Space(12) & Space(vEspacio + 4) _
                            & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Space(28 + vEspacio) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(5 + vEspacio) _
                                 & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO " & psOperacion, 25) & " ---- " & Space(9 + vEspacio) _
                            & " ---- " & ofun.ImpreFormat("EXTORNO " & psOperacion, 25) & " ---- " & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & Space(5 + vEspacio) _
                                & " Contrato  : " & lsCodAntiguo & Space(10) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & Space(5 + vEspacio) _
                                & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(10) & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & Space(vEspacio + 4) _
                            & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & Space(11 + vEspacio) _
                            & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & Space(10 + vEspacio) _
                            & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Space(6 + vEspacio) _
                            & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea
    Else
        lsCadImp = "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        'FRHU 20150817 OBSERVACION
        'lsCadImp = lsCadImp & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        If psOpeCod = gColRecOpePagTransfFocMacmEfe Or psOpeCod = gColRecOpePagTransfFocMacmVou Or psOpeCod = gColRecOpePagTransfFocMacmChq Then
            lsCadImp = lsCadImp & " Créditos Transf. FOCMAC" & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & " Pignoraticio-" & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
        End If
        'FIN FRHU 20150817
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO " & psOperacion, 25) & " ---- " & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
            lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        'lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea
    End If
    nPrintReciboExtorno = lsCadImp
    Set oFunI = Nothing
End Function

Public Function nPrintReciboPagoSobrante(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnSobrante As Currency, _
        ByVal psCodUser As String, ByVal psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim lsCadImp As String
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    Dim vEspacio As Integer
    vEspacio = 7
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Credito Pignoraticio " & Space(11) & Space(10 + vEspacio) _
                            & " Credito Pignoraticio " & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ofun.ImpreFormat(psNomAgencia, 24, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio) _
                            & " Contrato  : " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE:" & Trim(psNombCliente) & Space(vEspacio + 8) _
                            & " CLIENTE:" & Trim(psNombCliente) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & " --------------------------------------" & Space(5 + vEspacio) & " --------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Sobrante         : " & ofun.ImpreFormat(pnSobrante, 15) & Space(5 + vEspacio) _
                            & " Sobrante         : " & ofun.ImpreFormat(pnSobrante, 15) & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & " I.T.F.           : " &oFun.ImpreFormat(CDbl(pnITF), 15) & Space(5 + vEspacio) _
        '                    & " I.T.F.           : " &oFun.ImpreFormat(CDbl(pnITF), 15) & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & " Total Duplicado  : " &oFun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & Space(5 + vEspacio) _
        '                    & " Total Duplicado  : " &oFun.ImpreFormat(pnCostoDuplicado + CDbl(pnITF), 15) & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & " " & oFun.ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & Space(4 + vEspacio) _
        '                    & " " & oFun.ImpreCarEsp("Nro Duplicado    : ") & Format(pnNroDuplicado, "@@") & Space(17) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(39 + vEspacio) & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Sobrante de Remate ") & Space(23) & Space(vEspacio) _
                            & " " & ofun.ImpreCarEsp("Sobrante de Remate ") & Space(20) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " --------------------------------------" & Space(5 + vEspacio) _
                            & " --------------------------------------" & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 1) _
                            & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        nPrintReciboPagoSobrante = lsCadImp

Set oFunI = Nothing
End Function

Public Function PrintHojaInformativa(ByVal psTitulo As String, _
            Optional ByVal prPersonas As Recordset, _
            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, _
            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsCadImp As String
Dim lsDescLote As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
Dim lsCodPers As String * 13, lsApellido As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35


lnEspacio = 7

Dim i As Integer

Dim lsMateria As String * 14
Dim lsPesoBruto As String * 8
Dim lsPesoNeto As String * 9
Dim lsTasacion As String * 10
Dim lsDescripcion As String * 15


lsDescLote = lsDescLote & Chr$(13) & "Kilataje" & Space(6) & "P.Neto" & Space(2) & "P.Bruto" & Space(2) & "Tasación" & Space(2) & "Descripción"

'For i = 1 To FEJoyas.Rows - 1
'
'    lsMateria = Left(FEJoyas.TextMatrix(i, 1), 12)
'    lsPesoBruto = FEJoyas.TextMatrix(i, 2)
'    lsPesoNeto = FEJoyas.TextMatrix(i, 3)
'    lsTasacion = FEJoyas.TextMatrix(i, 4)
'    lsDescripcion = FEJoyas.TextMatrix(i, 5)
'
'    lsDescLote = lsDescLote & Chr$(13) & lsMateria & lsPesoBruto & _
'                 lsPesoNeto & lsTasacion & lsDescripcion
'
'Next

    
'===============Personas

lsDescLote = lsDescLote & Chr$(13)
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
'
        
prPersonas.MoveFirst
'If lnLongi = 35 Then
    lsCodPers = prPersonas!cPersCod
    lsApellido = prPersonas!cPersApellido
    lsNombre = prPersonas!cPersNombre
    'lsNombre = Left(prPersonas!cpersnombre, lnLongi)
    lsNroDocIdentidad = prPersonas!NroDNI
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
    lsTelefono = prPersonas!cPersTelefono
'Else
'    Do While Not prPersonas.EOF
'        If Right(lsNombre, 3) <> " / " Then
'            lsCodPers = prPersonas!cPersCod
'            lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
'            lsNroDocTributario = IIf(IsNull(prPersonas!NroRuc), "", prPersonas!NroRuc)
'            lsDireccion = IIf(IsNull(prPersonas!cPersDireccDomicilio), "", prPersonas!cPersDireccDomicilio)
'            lsZona = Mid(Trim(prPersonas!Zona), 1, 25)
'            lsCiudad = Mid(Trim(prPersonas!Prov), 1, 25) & " - " & Mid(Trim(prPersonas!Dpto), 1, 20)
'            lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
'        End If
'        lsNombre = lsNombre & Left(prPersonas!cpersnombre, lnLongi) & " / "
'        prPersonas.MoveNext
'    Loop
'    lsNombre = Left(lsNombre, Len(lsNombre) - 3)
'End If


lsCadImp = ""
lsCadImp = lsCadImp & Chr$(27) & Chr$(103)              'Tamaño  : 80, 77, 103
lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(20)
lsCadImp = lsCadImp & Chr$(27) & Chr$(65) & Chr$(9)     'Espaciado
        

    lsCadImp = lsCadImp & " " & String(11, oFunI.gPrnSaltoLinea)
    lsCadImp = lsCadImp & Space(8) & lsCodPers & Space(28) & Format(psFecContrato, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea  '12
    lsCadImp = lsCadImp & Space(49) & Format(psFecVencimiento, "dd/mm/yyyy")            '13
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
    lsCadImp = lsCadImp & Space(21) & psTitulo
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita  75

    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea   '15
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsApellido) & oFunI.gPrnSaltoLinea       '16
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsNombre) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea        '16
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea  ' Doc identid   19
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsDireccion) & " " & oFunI.gPrnSaltoLinea   ' Direccion 20
    lsCadImp = lsCadImp & Space(91) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea    'Plazo 21
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea   ' Zona    22
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea   ' Ciudad  23
    lsCadImp = lsCadImp & Space(16) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea   ' Telefono   24
    lsCadImp = lsCadImp & Space(16) & " " & oFunI.gPrnSaltoLinea   ' Telefono

    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

    lsCadImp = lsCadImp & Chr$(27) & Chr$(65) & oFunI.gPrnSaltoPagina      'Espaciado

    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnPiezas, 9, 2, True) & oFunI.gPrnSaltoLinea                    '27
    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnOroBruto, 9, 2, True) & oFunI.gPrnSaltoLinea                  '28
    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnOroNeto, 9, 2, True) & oFunI.gPrnSaltoLinea          '30
    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnValorTasacion, 9, 2, True) & oFunI.gPrnSaltoLinea    '31
    lsCadImp = lsCadImp & Space(76) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnMontoPrestado, 9, 2, True) & oFunI.gPrnSaltoLinea    '33
    lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(pnInteres, 9, 2, True) & oFunI.gPrnSaltoLinea                   '35
         

    lsCadImp = lsCadImp & Space(70) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 1) & " "
    lsCadImp = lsCadImp & Space(75 - 1 - Len(ofun.QuiebreTexto(lsDescLote, 1))) & ofun.ImpreFormat(pnCostoTasacion, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 2) & " "
    lsCadImp = lsCadImp & Space(75 - 1 - Len(ofun.QuiebreTexto(lsDescLote, 2))) & ofun.ImpreFormat(pnCostoCustodia, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 3) & " "
    lsCadImp = lsCadImp & Space(75 - 1 - Len(ofun.QuiebreTexto(lsDescLote, 3))) & ofun.ImpreFormat(pnImpuesto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 4) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 5) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 6) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 7) & " "
    lsCadImp = lsCadImp & Space(75 - 1 - Len(ofun.QuiebreTexto(lsDescLote, 7))) & ofun.ImpreFormat(pnMontoPrestado - pnInteres - pnCostoTasacion - pnCostoCustodia - pnImpuesto, 9, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 8) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 9) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 10) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 11) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 12) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 13) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 14) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 15) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 16) & " " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 17) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 17))) & "INC. 11: Pasados 30 dias del"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 18) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 18))) & " vencimiento  de su Contrato"
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea     'Desactiva Negrita
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 19) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 19))) & ofun.ImpreCarEsp(" sus Joyas entrarán a Remate")
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(1) & ofun.QuiebreTexto(lsDescLote, 20) & " "
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(72 - 25 - Len(ofun.QuiebreTexto(lsDescLote, 20))) & " sin Notificar."
    lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
    lsCadImp = lsCadImp & Space(1) & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "TASA EFECTIVA ANUAL: " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & oFunI.gPrnSaltoLinea
    If pnNroDuplicado > 0 Then
        lsCadImp = lsCadImp & Space(13) & "Duplicado Nro : " & ofun.ImpreFormat(pnNroDuplicado, 5, 0) & oFunI.gPrnSaltoLinea
    End If
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(13) & "" & String(5, oFunI.gPrnSaltoLinea)


    lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    lsCadImp = lsCadImp & Space(45) & psCodUser
    lsCadImp = lsCadImp & Space(34) & psTitulo
    lsCadImp = lsCadImp & Space(10) & "" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(96) & Str(pnPlazo)

    
    PrintHojaInformativa = lsCadImp
    Set oFunI = Nothing
End Function


Public Function ImprimirRescate(ByVal psCtaCod As String, ByVal psCliente As String, ByVal pnprestamo As Double, _
                                ByVal pnSaldo As Double, pdfecha As Date, psUser As String, Optional psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim vEspacio As Integer

Dim ofun As New COMFunciones.FCOMImpresion
Dim oCon As New COMConecta.DCOMConecta
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
oCon.AbreConexion
    vEspacio = 7
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "       DEVOLUCION DE JOYAS " & Space(11) & Space(8 + vEspacio) _
                            & " DEVOLUCION DE JOYAS " & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat("CMAC MAYNAS", 24, 0) & Space(2) & Format(pdfecha & " " & oCon.GetHoraServer(), "dd/mm/yyyy hh:mm") & Space(4) _
                            & " " & ofun.ImpreFormat("CMAC MAYNAS", 24, 0) & Space(2) & Format(pdfecha & " " & oCon.GetHoraServer(), "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(8) & Space(vEspacio) _
                            & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(11) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE:" & Trim(psCliente) & Space(vEspacio + 5) _
                            & " CLIENTE:" & Trim(psCliente) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & " ------------------------------------------" & Space(4) & " ------------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Préstamo       : " & ofun.ImpreFormat(pnprestamo, 15) & Space(5 + vEspacio) _
                            & " Préstamo       : " & ofun.ImpreFormat(pnprestamo, 15) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Saldo Capital  : " & ofun.ImpreFormat(pnSaldo, 15) & Space(5 + vEspacio) _
                            & " Saldo Capital  : " & ofun.ImpreFormat(pnSaldo, 15) & oFunI.gPrnSaltoLinea
       
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psUser, "@@@@") & Space(36 + vEspacio) & Format(psUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " ------------------------------------------" & Space(4) _
                            & " ------------------------------------------" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                            
'        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psMensaje, 40) & Space(vespacio + 1) _
'                            & "" & ofun.ImpreFormat(psMensaje, 40) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        ImprimirRescate = lsCadImp
Set oFunI = Nothing
End Function

Public Function ImpBloqueo(ByVal pBloqueo As Boolean, ByVal pCodCta As String, ByVal pPlazo As String, ByVal pNombre As String, ByVal gsCodUser As String, ByVal gsNomAge As String, ByVal gdFecSis As Date, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    'Dim vNombre As String * 28
    Dim vEspacio As Integer
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    vEspacio = 9
    
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & Space(15) & " CREDITO PIGNORATICIO" & Space(22 + vEspacio) & "CREDITO PIGNORATICIO" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & oFunI.gPrnSaltoLinea    'Desactiva Negrita
        lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vEspacio) & ofun.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & Space(4) & " CONTRATO(" & Format(pPlazo, "00") & "):" & ofun.ImpreFormat(pCodCta, 12, 17) & Space(vEspacio) & "CONTRATO(" & Format(pPlazo, "00") & "):" & ofun.ImpreFormat(pCodCta, 12, 17) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Space(4) & " CLIENTE     : " & ofun.ImpreFormat(pNombre, 28, 0) & Space(vEspacio) & "CLIENTE     : " & ofun.ImpreFormat(pNombre, 28, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        If pBloqueo Then
            lsCadImp = lsCadImp & Space(4) & " BLOQUEO DE CONTRATO" & Space(vEspacio + 23) & "BLOQUEO DE CONTRATO" & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & Space(4) & " DESBLOQUEO DE CONTRATO" & Space(vEspacio + 20) & "DESBLOQUEO DE CONTRATO" & oFunI.gPrnSaltoLinea
        End If
  
        lsCadImp = lsCadImp & Space(4) & Space(22 + 16) & Format(gsCodUser, "@@@@") & Space(vEspacio) & Space(16 + 22) & Format(gsCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        ImpBloqueo = lsCadImp
    Set oFunI = Nothing
End Function

Public Function ImprimePolizaNueva(ByVal psNomAdj As String, ByVal psNatAdj As String, ByVal psTriAdj As String, ByVal psVarNroRemate As String, ByVal pdFecSis As Date, ByVal psSubTotalBoleta As String, ByVal psIGV As String, ByVal psPreVentaNeto As String, Rs As ADODB.Recordset, rs1 As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    Dim lsCadImp As String
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim oFunC As New COMFunciones.FCOMCadenas
    Dim i As Integer
    Dim nLinea As Integer
    Dim nTLinea As Integer
    oFunI.Inicia psImpresora
    lsCadImp = ""
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(18) & Trim(psNomAdj) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(18) & IIf(Trim(psNatAdj) = "", Trim(psTriAdj), Trim(psNatAdj)) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
  
    'Cliente que joyas son rematadas
    If Not (Rs.EOF And Rs.BOF) Then
        Rs.MoveFirst
        lsCadImp = lsCadImp & Space(18) & Trim(Rs(0)) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(18) & IIf(Trim(Rs(3)) = "", Trim(Rs(4)), Trim(Rs(3))) & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & Space(75 - Len(psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy"))) & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
    End If
    lsCadImp = lsCadImp & Space(24) & "Portal Espinar Nro 146 " & Space(60) & pdFecSis & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    i = 0
    If Not (rs1.EOF And rs1.BOF) Then
        nLinea = rs1.RecordCount
        rs1.MoveFirst
        Do Until rs1.EOF
            If i = 0 Then
                lsCadImp = lsCadImp & Space(5) & Trim(rs1(0)) & Space(15) & ofun.ImpreFormat(Trim(rs1(1)), 45) & Space(70) & psPreVentaNeto & oFunI.gPrnSaltoLinea
            Else
                lsCadImp = lsCadImp & Space(5) & Trim(rs1(0)) & Space(15) & ofun.ImpreFormat(Trim(rs1(1)), 45) & Space(70) & oFunI.gPrnSaltoLinea
            End If
            i = i + 1
            rs1.MoveNext
        Loop
        
    End If
    nTLinea = 11
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    For i = 1 To nTLinea - nLinea
       lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    Next

    
    lsCadImp = lsCadImp & Space(10) & oFunC.UnNumero(psPreVentaNeto) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(140) & Format(psPreVentaNeto, "#0.00") & oFunI.gPrnSaltoLinea

    ImprimePolizaNueva = lsCadImp
    Set oFunI = Nothing
End Function

'Public Function ImprimePolizaNueva(ByVal psNomAdj As String, ByVal psNatAdj As String, ByVal psTriAdj As String, ByVal psVarNroRemate As String, ByVal pdFecSis As Date, ByVal psSubTotalBoleta As String, ByVal psIGV As String, ByVal psPreVentaNeto As String, rs As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
'    Dim lsCadImp As String
'    Dim oFunI As New COMFunciones.FCOMVarImpresion
'    oFunI.Inicia psImpresora
'
'    lsCadImp = lsCadImp & ""
'    lsCadImp = lsCadImp & "Adjudicatario:" & Trim(psNomAdj) & Space(75 - Len("Adjudicatario:" & Trim(psNomAdj))) & "Adjudicatario:" & Trim(psNomAdj) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj) & Space(75 - Len("DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj))) & "DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj) & oFunI.gPrnSaltoLinea
'
'    'Cliente que joyas son rematadas
'    If Not (rs.EOF And rs.BOF) Then
'        rs.MoveFirst
'        lsCadImp = lsCadImp & "Ejecutado:   " & Trim(rs(0)) & Space(75 - Len("Ejecutado:   " & Trim(rs(0)))) & "Ejecutado:   " & Trim(rs(0)) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & "DNI: " & Trim(rs(3)) & " " & Trim(rs(4)) & Space(75 - Len("DNI: " & Trim(rs(3)) & " " & Trim(rs(4)))) & "DNI: " & Trim(rs(3)) & " " & Trim(rs(4)) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & Space(75 - Len(psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy"))) & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
'    End If
'
'    lsCadImp = lsCadImp & String(45, "-") & Space(75 - Len(String(45, "-"))) & String(45, "-") & oFunI.gPrnSaltoLinea
'
'        If Not (rs.EOF And rs.BOF) Then
'            rs.MoveFirst
'
'            lsCadImp = lsCadImp & " " & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & " " & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & "Adjudicatario:" & Trim(psNomAdj) & Space(75 - Len("Adjudicatario:" & Trim(psNomAdj))) & "Adjudicatario:" & Trim(psNomAdj) & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & "DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj) & Space(75 - Len("DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj))) & "DNI: " & Trim(psNatAdj) & " RUC: " & Trim(psTriAdj) & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & " "
'                'Cliente que joyas son rematadas
'
'            lsCadImp = lsCadImp & "Ejecutado:   " & Trim(rs(0)) & Space(75 - Len("Ejecutado:   " & Trim(rs(0)))) & "Ejecutado:   " & Trim(rs(0)) & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & "DNI: " & Trim(rs(3)) & " " & Trim(rs(4)) & Space(75 - Len("DNI: " & Trim(rs(3)) & " " & Trim(rs(4)))) & "DNI: " & Trim(rs(3)) & " " & Trim(rs(4)) & oFunI.gPrnSaltoLinea
'            lsCadImp = lsCadImp & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & Space(75 - Len(psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy"))) & psVarNroRemate & Space(18) & Format(pdFecSis, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
'           '
'            Do Until rs.EOF
'                lsCadImp = lsCadImp & String(45, "-") & Space(75 - Len(String(45, "-"))) & String(45, "-") & oFunI.gPrnSaltoLinea
'                lsCadImp = lsCadImp & rs(0) & Space(1) & Left(rs(2), 26) & rs(1) & " gr" & Space(75 - Len(rs(0) & Space(1) & Left(rs(2), 26) & rs(1) & " gr")) & rs(0) & Space(1) & Left(rs(2), 26) & rs(1) & " gr" & oFunI.gPrnSaltoLinea
'                rs.MoveNext
'            Loop
'
'        End If
'
'    lsCadImp = lsCadImp & ""
'    lsCadImp = lsCadImp & "TOTAL BOLETA: " & Format(psSubTotalBoleta, "#0.00") & Space(75 - Len("TOTAL BOLETA: " & Format(psSubTotalBoleta, "#0.00"))) & Space(5) & "TOTAL BOLETA: " & Format(psSubTotalBoleta, "#0.00") & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "ITF         : " & Format(psIGV, "#0.00") & Space(75 - Len("I.G.V.      : " & Format(psIGV, "#0.00"))) & Space(7) & "ITF      : " & Format(psIGV, "#0.00") & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "VENTA NETA  : " & Format(psPreVentaNeto, "#0.00") & Space(75 - Len("VENTA NETA  : " & Format(psPreVentaNeto, "#0.00"))) & Space(5) & "VENTA NETA  : " & Format(psPreVentaNeto, "#0.00") & oFunI.gPrnSaltoLinea
'
'    ImprimePolizaNueva = lsCadImp
'    Set oFunI = Nothing
'End Function


'Procedimiento de impresión del recibo de venta en subasta
Public Function ImpRecVtaSub(pCodCta As String, pNomAdj As String, pMonCom As Currency, ByVal gsNomAge As String, _
                            ByVal gdFecSis As Date, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    Dim vNombre As String * 28
    Dim vEspacio As Integer
    Dim sCadImp As String
    Dim oFunI As New COMFunciones.FCOMImpresion
    Dim ofun As New COMFunciones.FCOMVarImpresion
    ofun.Inicia psImpresora
  
    vNombre = pNomAdj
    vEspacio = 6
  
    sCadImp = sCadImp & Chr$(27) & Chr$(77)
    sCadImp = sCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    sCadImp = sCadImp & Space(5) & " CREDITO PIGNORATICIO" & Space(35) & "CREDITO PIGNORATICIO" & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & ""
    sCadImp = sCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    sCadImp = sCadImp & oFunI.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(10) & oFunI.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
    sCadImp = sCadImp & "CONTRATO     :" & Space(8) & pCodCta & Space(12) & "CONTRATO       :" & Space(8) & pCodCta & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & "ADJUDICATARIO  :" & Space(3) & Trim(vNombre) & Space(13) & "ADJUDICATARIO  :" & Space(3) & Trim(vNombre) & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & ""
    sCadImp = sCadImp & "MONTO DE VENTA :" & Space(3) & oFunI.ImpreFormat(pMonCom, 20, , True) & Space(12) & "MONTO DE VENTA :" & Space(3) & oFunI.ImpreFormat(pMonCom, 20, , True) & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & ""
    sCadImp = sCadImp & "Venta en Subasta  " & Space(20) & Format(gsCodUser, "@@@@") & Space(12) & "Venta en Subasta   " & Space(19) & Format(gsCodUser, "@@@@") & ofun.gPrnSaltoLinea
    sCadImp = sCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    ImpRecVtaSub = sCadImp
    Set oFunI = Nothing
    
End Function


'Procedimiento de impresión del recibo de venta en subasta
Public Function ImpRecupSub(ByVal pCodCta As String, ByVal pNomAdj As String, ByVal pMonRecup As Currency, ByVal pnTotRecup As Currency, ByVal pnMonITF As Currency, ByVal gsNomAge As String, _
                            ByVal gdFecSis As Date, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON, _
                            Optional ByVal pbImpTMU As Boolean = False, _
                            Optional ByVal pnDsctoint As Double = 0) As String
    Dim vNombre As String * 28
    Dim vEspacio As Integer
    Dim sCadImp As String
    Dim oFunI As New COMFunciones.FCOMImpresion
    Dim ofun As New COMFunciones.FCOMVarImpresion
    
    ofun.Inicia psImpresora
  
    vNombre = pNomAdj
    vEspacio = 10
    If Not pbImpTMU Then
        sCadImp = sCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        sCadImp = sCadImp & Space(10) & " CREDITO PIGNORATICIO" & Space(35) & "CREDITO PIGNORATICIO" & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & ""
        sCadImp = sCadImp & Chr$(27) & Chr$(70)      'Desactiva Negrita
        sCadImp = sCadImp & oFunI.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vEspacio) & oFunI.ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & ofun.gPrnSaltoLinea & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        sCadImp = sCadImp & "CONTRATO     :" & Space(8) & pCodCta & Space(vEspacio) & "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & "CLIENTE      :" & vNombre & Space(vEspacio) & "CLIENTE      :" & vNombre & ofun.gPrnSaltoLinea 'PASI20150930 agrego el Salto de Linea
        'Print #ArcSal, ""
        sCadImp = sCadImp & "MONTO DE RECUPERACION:" & oFunI.ImpreFormat(pMonRecup + pnDsctoint, 15, , True) & Space(vEspacio + 2) & "MONTO DE RECUPERACION:" & oFunI.ImpreFormat(pMonRecup + pnDsctoint, 15, , True) & ofun.gPrnSaltoLinea
        'PEAC 20200904
        sCadImp = sCadImp & "DESCUENTO            :" & oFunI.ImpreFormat(pnDsctoint, 15, , True) & Space(vEspacio + 2) & "DESCUENTO            :" & oFunI.ImpreFormat(pnDsctoint, 15, , True) & ofun.gPrnSaltoLinea
        
        sCadImp = sCadImp & "ITF EFECTIVO         :" & oFunI.ImpreFormat(pnMonITF, 15, , True) & Space(vEspacio + 2) & "ITF EFECTIVO         :" & oFunI.ImpreFormat(pnMonITF, 15, , True) & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & "TOTAL DE RECUPERACION:" & oFunI.ImpreFormat(pnTotRecup, 15, , True) & Space(vEspacio + 2) & "TOTAL DE RECUPERACION:" & oFunI.ImpreFormat(pnTotRecup, 15, , True) & ofun.gPrnSaltoLinea
      '  Print #ArcSal, ""
        sCadImp = sCadImp & "JOYA(S) ADJUDICADA RECUPERADA" & Space(8) & Format(gsCodUser, "@@@@") & Space(vEspacio + 1) & "JOYA(S)ADJUDICADA RECUPERADA" & Space(8) & Format(gsCodUser, "@@@@") & ofun.gPrnSaltoLinea
        
        sCadImp = sCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
    Else
        'sCadImp = sCadImp & " " & Chr$(27) & Chr$(69)    'Activa Negrita
        sCadImp = sCadImp & " " & Space(10) & " CREDITO PIGNORATICIO" & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & " "
        sCadImp = sCadImp & Chr$(27) & Chr$(70)      'Desactiva Negrita
        sCadImp = sCadImp & gsNomAge & " -- " & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & ofun.gPrnSaltoLinea
        'sCadImp = sCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        sCadImp = sCadImp & "CONTRATO     :" & Space(8) & pCodCta & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & "CLIENTE      :" & vNombre & ofun.gPrnSaltoLinea
        'Print #ArcSal, ""
        sCadImp = sCadImp & "MONTO DE RECUPERACION:" & oFunI.ImpreFormat(pMonRecup + pnDsctoint, 15, , True) & ofun.gPrnSaltoLinea
        
        'PEAC 20200904
        sCadImp = sCadImp & "DESCUENTO            :" & oFunI.ImpreFormat(pnDsctoint, 15, , True) & ofun.gPrnSaltoLinea
        
        sCadImp = sCadImp & "ITF EFECTIVO         :" & oFunI.ImpreFormat(pnMonITF, 15, , True) & ofun.gPrnSaltoLinea
        sCadImp = sCadImp & "TOTAL DE RECUPERACION:" & oFunI.ImpreFormat(pnTotRecup, 15, , True) & ofun.gPrnSaltoLinea
      '  Print #ArcSal, ""
        sCadImp = sCadImp & "JOYA(S) ADJUDICADA RECUPERADA" & Space(5) & Format(gsCodUser, "@@@@") & ofun.gPrnSaltoLinea
        
        sCadImp = sCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    End If
        
        
        ImpRecupSub = sCadImp
    Set oFunI = Nothing
End Function

'*** PEAC 20090703
Public Function nPrintReciboExtorAboCta(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psCtaAhoExt As String, ByVal pnITFAho As Currency, ByVal psNombCliente As String, ByVal psOperacion As String, _
        ByVal pnMonto As Currency, ByVal pnSaldoCapAnt As Currency, ByVal pnMovExtornado As Long, _
        ByVal psCodUser As String, ByVal psVarPersCMAC As String, _
        Optional ByVal psmensaje As String = "", Optional ByVal psUsuOri As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False) As String
        
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim loConstSistema As New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
        
    Dim vEspacio As Integer
    vEspacio = 28
        
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    If pbImpTMU = False Then
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & Space(12) & Space(vEspacio + 4) _
                            & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Space(28 + vEspacio) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(5 + vEspacio) _
                                 & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & Space(vEspacio + 4) _
                            & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & Space(5 + vEspacio) _
                            & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO DE ABONO", 25) & " ---- " & Space(9 + vEspacio) _
                            & " ---- " & ofun.ImpreFormat("EXTORNO DE ABONO", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & Space(11 + vEspacio) _
                            & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & Space(11 + vEspacio) _
                            & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & Space(10 + vEspacio) _
                            & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Space(10 + vEspacio) _
                            & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Else
        lsCadImp = "" & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO DE ABONO", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)

    End If
    nPrintReciboExtorAboCta = lsCadImp
    Set oFunI = Nothing
End Function

'MAVM 20100717
Public Function ImprimeVolanteTasacionJoyas(pnOroBruto As Double, pnOroNeto As Double, pnValorTasacion As Currency, pnPiezas As Integer, prjoyas As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim lnLineas As Integer, lnCont As Integer, i As Integer
    oFunI.Inicia psImpresora
    lsCadImp = ""
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Do While Not prjoyas.EOF
        lnCont = lnCont + 1
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Pzas"), 5, 0)
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Descripcion"), 33, 0)
        lsCadImp = lsCadImp & Space(4) & ofun.ImpreFormat(Right(prjoyas("Material"), 2), 10, 0)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PNeto"), 8, 2)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PBruto"), 8, 2) & oFunI.gPrnSaltoLinea
        If lnCont > 10 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(20) + "Continua siguiente pagina..."
            lnCont = 0
            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        End If
        prjoyas.MoveNext
    Loop
    lnLineas = 14 - lnCont
    For i = 1 To lnLineas
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    Next
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(6) & ofun.ImpreFormat(pnPiezas, 5, 0) & Space(4) & ofun.ImpreFormat(pnOroBruto, 9, 2) & Space(25) & ofun.ImpreFormat(pnOroNeto, 9, 2) & Space(3) & ofun.ImpreFormat(pnValorTasacion, 9, 2) & oFunI.gPrnSaltoLinea
    ImprimeVolanteTasacionJoyas = lsCadImp
    Set oFunI = Nothing
    Set ofun = Nothing
End Function

'***MAVM 20100717
Public Function nPrintRegJoyaGarantia(ByVal psJoyGarCod As String, Optional ByVal prPersonas As ADODB.Recordset, _
            Optional ByVal psFecContrato As String, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
            Optional ByVal pnValorTasacion As Currency, Optional ByVal pnPiezas As Double, _
            Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsCadImp As String, lsSM As String
Dim lsDescLote As String
Dim lnEspacio As Integer
Dim lnLongi As Integer, lnX As Integer
Dim lsApellido As String, lsNombreLargo As String
Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
Dim lmDetalle(20) As String
Dim pnInteresDif As Double, pnMontoRec As Double
Dim ofun As New COMFunciones.FCOMImpresion
Dim cTempoTexto75 As String * 75
Dim lnSum14k As Double
Dim lnSum16k As Double
Dim lnSum18k As Double
Dim lnSum21k As Double
Dim oFunI As New COMFunciones.FCOMVarImpresion

oFunI.Inicia psImpresora

    Dim lrCredPigJoyasDet As New ADODB.Recordset
    Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato

    Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosRegJoyasDet(psJoyGarCod)
    Set loMuestraContrato = Nothing

        Dim item As Integer, itemTMP As Integer
        item = 0
        Do While Not lrCredPigJoyasDet.EOF
            lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescripcion, 31) & " " _
                        & lrCredPigJoyasDet!cKilataje & "K " & Space(16) & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
            If Trim(lrCredPigJoyasDet!cKilataje) = 14 Then
                lnSum14k = lnSum14k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 16 Then
                lnSum16k = lnSum16k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 18 Then
                lnSum18k = lnSum18k + lrCredPigJoyasDet!npesoneto
            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 21 Then
                lnSum21k = lnSum21k + lrCredPigJoyasDet!npesoneto
            End If
            item = item + 1
            lrCredPigJoyasDet.MoveNext
        Loop
        For itemTMP = item To 20
            lmDetalle(itemTMP) = ""
        Next
        lrCredPigJoyasDet.Close

lnEspacio = 7
'''lsSM = "S/. " 'marg ers044-2016
lsSM = gcPEN.gcPEN_SIMBOLO & " " 'marg ers044-2016
lsCadImp = ""
lsDescLote = lsDescLote & Chr$(13)
prPersonas.MoveFirst

Dim sSql As String, RSTEMP As ADODB.Recordset, cn As COMConecta.DCOMConecta
Set RSTEMP = New ADODB.Recordset
Set cn = New COMConecta.DCOMConecta
sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre,cPersNombre as cPersNombreLargo "
sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
If RSTEMP.State = 1 Then RSTEMP.Close
    cn.AbreConexion
    Set RSTEMP = cn.CargaRecordSet(sSql)
    If RSTEMP.State = 1 Then
    If RSTEMP.RecordCount > 0 Then
        lsApellido = RSTEMP!cPersApellido
        lsNombre = RSTEMP!cPersNombre
        lsNombreLargo = RSTEMP!cPersNombreLargo
    End If
    If RSTEMP.State = 1 Then RSTEMP.Close
End If
cn.CierraConexion
Set cn = Nothing

    lsCodPers = prPersonas!cPersCod
    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
    lsDireccion = prPersonas!cPersDireccDomicilio
    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
    lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
    
    lsCadImp = ""
    
    '**Configuración de página**********************************************************
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    '***********************************************************************************
    
    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 8)
    lsCadImp = lsCadImp & Space(20) & lsCodPers & Space(21) & Format(psFecContrato, "dd/MM/yyyy") & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(54) & Format(psFecVencimiento, "dd/MM/yyyy") & Space(14)
    lsCadImp = lsCadImp & Space(54) & Space(22)

    lsCadImp = lsCadImp & oFunI.gPrnBoldON & psJoyGarCod & oFunI.gPrnBoldOFF & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNombreLargo) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(26) & ofun.ImpreFormat(ofun.ImpreCarEsp(lsDireccion), 50, 0) & Space(7) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreFormat(ofun.ImpreCarEsp(lsDireccion), 50, 0) & Space(9) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnPiezas, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoPrestado, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnPiezas, 12, 2, True) & Space(35) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroBruto, 12, 2, True) & Space(34) & ofun.ImpreFormat(pnTasaInteresMensual, 12, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroBruto, 12, 2, True) & Space(36) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroNeto, 12, 2, True) & Space(30) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    'lsCadImp = lsCadImp & Space(26) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoRec, 12, 2, True) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(26) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & Space(30) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
    'Titulo y Monto Prestamo
    cTempoTexto75 = ""
    'Primer Item y el Prestamo
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(0), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnMontoPrestado, 18, 2, True) & oFunI.gPrnSaltoLinea
    'Segundo Item y interes
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(1), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnInteresDif, 18, 2, True) & oFunI.gPrnSaltoLinea
    'tercer Item y Costo de Tasacion
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(2), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoTasacion, 18, 2, True) & oFunI.gPrnSaltoLinea  '4
    'Cuarto Item y Costo de Custodia
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(3), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoCustodia, 18, 2, True) & oFunI.gPrnSaltoLinea  '5
    'Quinto Item y el Impuesto
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(4), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea  '6
    'Item 6
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(5), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    'Item 7
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(6), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    'Item 8
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(7), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    'Item 9
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(8), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & Space(22) & oFunI.gPrnSaltoLinea
    'Item 10 y Neto a Recibir
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(9), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea

    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(10), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(11), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(12), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(13), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(14), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(15), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(16), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(17), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(18), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
    cTempoTexto75 = ""
    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(19), 75)
    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoPagina '& oFunI.gPrnSaltoLinea
    nPrintRegJoyaGarantia = lsCadImp
    Set oFunI = Nothing
End Function

Public Function ImprimirRescateJoyGar(ByVal psCtaCod As String, ByVal psCliente As String, ByVal pnprestamo As Double, _
                                ByVal pnSaldo As Double, pdfecha As Date, psUser As String, Optional psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim vEspacio As Integer

Dim ofun As New COMFunciones.FCOMImpresion
Dim oCon As New COMConecta.DCOMConecta
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
oCon.AbreConexion
    vEspacio = 7
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & "       DEVOLUCION DE JOYAS " & Space(11) & Space(8 + vEspacio) _
                            & " DEVOLUCION DE JOYAS " & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat("CMAC MAYNAS", 24, 0) & Space(2) & Format(pdfecha & " " & oCon.GetHoraServer(), "dd/mm/yyyy hh:mm") & Space(4) _
                            & " " & ofun.ImpreFormat("CMAC MAYNAS", 24, 0) & Space(2) & Format(pdfecha & " " & oCon.GetHoraServer(), "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(8) & Space(vEspacio) _
                            & " Contrato  : " & Mid(psCtaCod, 1, 5) & "-" & Mid(psCtaCod, 6) & Space(11) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE:" & Trim(psCliente) & Space(vEspacio + 5) _
                            & " CLIENTE:" & Trim(psCliente) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)           'Activa Negrita
        lsCadImp = lsCadImp & " ------------------------------------------" & Space(4) & " ------------------------------------------" & oFunI.gPrnSaltoLinea
        
        
'        lsCadImp = lsCadImp & " Préstamo       : " & ofun.ImpreFormat(pnprestamo, 15) & Space(5 + vEspacio) _
'                            & " Préstamo       : " & ofun.ImpreFormat(pnprestamo, 15) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & " Saldo Capital  : " & ofun.ImpreFormat(pnsaldo, 15) & Space(5 + vEspacio) _
'                            & " Saldo Capital  : " & ofun.ImpreFormat(pnsaldo, 15) & oFunI.gPrnSaltoLinea
       
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psUser, "@@@@") & Space(36 + vEspacio) & Format(psUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " ------------------------------------------" & Space(4) _
                            & " ------------------------------------------" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                            
'        lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psMensaje, 40) & Space(vespacio + 1) _
'                            & "" & ofun.ImpreFormat(psMensaje, 40) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        ImprimirRescateJoyGar = lsCadImp
Set oFunI = Nothing
End Function
'RECO20141203 ERS163-2014**************************************
Public Function nPrintVolanteTasacionCamp(pnPlazo As Integer, pnOroBruto As Double, pnOroNeto As Double, pnValorTasacion As Currency, pnPiezas As Integer, prjoyas As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim lnLineas As Integer, lnCont As Integer, i As Integer
    oFunI.Inicia psImpresora
    lsCadImp = ""
    '**Configuración de página**********************************************************
    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
    '***********************************************************************************
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnPlazo, 3, 0) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Do While Not prjoyas.EOF
        lnCont = lnCont + 1
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Pzas"), 5, 0)
        lsCadImp = lsCadImp & Space(7) & ofun.ImpreFormat(prjoyas("Descripcion"), 33, 0)
        lsCadImp = lsCadImp & Space(4) & ofun.ImpreFormat(Right(prjoyas("Material"), 2), 10, 0)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PNeto"), 8, 2)
        lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(prjoyas("PBruto"), 8, 2) & oFunI.gPrnSaltoLinea
        If lnCont > 10 Then
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(20) + "Continua siguiente pagina..."
            lnCont = 0
            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(72) & ofun.ImpreFormat(pnPlazo, 3, 0) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        End If
        prjoyas.MoveNext
    Loop
    prjoyas.MoveFirst
    lnLineas = 14 - lnCont
    For i = 1 To lnLineas
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    Next
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(6) & ofun.ImpreFormat(pnPiezas, 5, 0) & Space(4) & ofun.ImpreFormat(pnOroBruto, 9, 2) & Space(25) & ofun.ImpreFormat(pnOroNeto, 9, 2) & Space(3) & ofun.ImpreFormat(pnValorTasacion, 9, 2) & oFunI.gPrnSaltoLinea
    nPrintVolanteTasacionCamp = lsCadImp
    Set oFunI = Nothing
    Set ofun = Nothing
End Function
'RECO FIN******************************************************

'*** PEAC 20161129
'RIRO campana prendario
Public Function RecuperaDatosHojaResumenPigno(ByVal psCtaCod As String, _
                ByRef prCredPig As ADODB.Recordset, _
                ByRef prCredPigJoyas As ADODB.Recordset, _
                ByRef prCredPigPersonas As ADODB.Recordset, _
                ByRef prCredPigCostos As ADODB.Recordset, _
                ByRef prCredPigJoyasDet As ADODB.Recordset, _
                ByRef prPigTasas As ADODB.Recordset, _
                ByRef prPigCosNot As ADODB.Recordset, _
                Optional ByRef nCampana As Integer = 0 _
                )

'stp_sel_ObtieneCosNotarialPigno rsPigCosNot
    Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    
    Set prCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCtaCod)
    Set prCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCtaCod)
    Set prCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCtaCod)
    Set prCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCtaCod)
    Set prCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCtaCod)
    Set prPigTasas = loMuestraContrato.dObtieneDatosCreditoPignoraticioTasas(psCtaCod)
    Set prPigCosNot = loMuestraContrato.nObtieneDatosObtieneCosNotarialPignoraticio(psCtaCod)
    nCampana = loMuestraContrato.CampDevuelveCampana(psCtaCod)
    
    Set loMuestraContrato = Nothing
    
End Function
'RIRO campana prendario

''*** PEAC 20161128
'Public Function nImpreHojaResumenCredGaranMobiliariaOro(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
'            Optional ByVal prPersonas As ADODB.Recordset, _
'            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
'            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
'            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
'            Optional ByVal pnValorTasacion As Currency, _
'            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
'            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
'            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
'            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pnITF As Double) As String
'
'' Imprime Hoja Resumen Credito con carantia mobiliaria de oro en Formato PDF
''Public Function nImpreHojaResumenCredGaranMobiliariaOro(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, _
'            Optional ByVal prPersonas As ADODB.Recordset, _
'            Optional ByVal pnTasaInteresMensual As Double, Optional ByVal pnMontoPrestado As Currency, _
'            Optional ByVal psFecContrato As String, Optional ByVal psFecVencimiento As String, _
'            Optional ByVal pnPlazo As Integer, Optional ByVal pnOroBruto As Double, Optional ByVal pnOroNeto As Double, _
'            Optional ByVal pnValorTasacion As Currency, _
'            Optional ByVal pnPiezas As Double, Optional ByVal psLote As String, _
'            Optional ByVal pn14K As Double, Optional ByVal pn16K As Double, Optional ByVal pn18K As Double, Optional ByVal pn21K As Double, _
'            Optional ByVal pnInteres As Currency, Optional ByVal pnCostoTasacion As Currency, Optional ByVal pnCostoCustodia As Currency, _
'            Optional ByVal pnImpuesto As Currency, Optional ByVal psCodUser As String, Optional ByVal pnNroDuplicado As Integer = 0, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pnITF As Double) As String
'
'
'Dim lsCadImp As String, lsSM As String
'
'Dim lsDescLote As String
'
'Dim lnEspacio As Integer
'Dim lnLongi As Integer, lnX As Integer
'
'Dim lsApellido As String, lsNombreLargo As String
'Dim lsCodPers As String, lsNombre As String, lsNroDocIdentidad As String, lsNroDocTributario As String
'Dim lszona As String * 35, lsCiudad As String * 35, lsDireccion As String * 35, lsTelefono As String * 35
'Dim lmDetalle(20) As String
'Dim pnInteresDif As Double, pnMontoRec As Double
'Dim ofun As New COMFunciones.FCOMImpresion
'Dim lnTicTEA As Double, lnTasInt As Double '*** PEAC 20101220
''lnTicTEA = Round((((1 + lnTasInt / 100) ^ 12 - 1) * 100), 2)
'
'Dim cTempoTexto75 As String * 75
'
'Dim lnSum14k As Double
'Dim lnSum16k As Double
'Dim lnSum18k As Double
'Dim lnSum21k As Double
'
'Dim oFunI As New COMFunciones.FCOMVarImpresion
'oFunI.Inicia psImpresora
'
'pnInteresDif = pnInteres
'
'If pbBuscaData = True Then
'
'    Dim lrCredPig As New ADODB.Recordset
'    Dim lrCredPigJoyas As New ADODB.Recordset
'    Dim lrCredPigPersonas As New ADODB.Recordset
'    Dim lrCredPigCostos As New ADODB.Recordset
'    Dim lrCredPigJoyasDet As New ADODB.Recordset
'    Dim lrPigTasas As New ADODB.Recordset
'
'    Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato
'
'    Set lrPigTasas = loMuestraContrato.dObtieneDatosCreditoPignoraticioTasas(psCodCta)
'    Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCodCta)
'    Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCodCta)
'    Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
'    Set lrCredPigCostos = loMuestraContrato.nObtieneDatosCostosRegistroCredPignoraticio(psCodCta)
'    Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCodCta)
'    Set loMuestraContrato = Nothing
'
'    If lrCredPig.BOF And lrCredPig.EOF Then
'        Set lrCredPig = Nothing
'        Set lrCredPigJoyas = Nothing
'        Set lrCredPigPersonas = Nothing
'        psmensaje = " No se encuentra el Credito Pignoraticio " & psCodCta
'        Exit Function
'    Else
'
'        lnTasInt = lrCredPig!nTasaInteres
'
'        pnOroBruto = lrCredPig!nOroBruto
'        pnOroNeto = lrCredPig!nOroNeto
'        pnPiezas = lrCredPig!nPiezas
'        pnValorTasacion = lrCredPig!nTasacion
'        pnMontoPrestado = lrCredPig!nMontoCol
'        pnInteresDif = IIf(IsNull(lrCredPig!nIntDif), pnInteres, lrCredPig!nIntDif)
'
'        pnMontoRec = pnMontoPrestado
'
'        pnPlazo = IIf(IsNull(lrCredPig!nPlazo) = True, 0, lrCredPig!nPlazo)
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
'
'        pnImpuesto = pnITF
'
'        lrCredPigCostos.Close
'        Set lrCredPigCostos = Nothing
'        Dim item As Integer, itemTMP As Integer
'        item = 0
'        Do While Not lrCredPigJoyasDet.EOF
'            lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 31) & " " _
'                        & lrCredPigJoyasDet!cKilataje & "K " & Space(16) & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
'            If Trim(lrCredPigJoyasDet!cKilataje) = 14 Then
'                lnSum14k = lnSum14k + lrCredPigJoyasDet!npesoneto
'            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 16 Then
'                lnSum16k = lnSum16k + lrCredPigJoyasDet!npesoneto
'            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 18 Then
'                lnSum18k = lnSum18k + lrCredPigJoyasDet!npesoneto
'            ElseIf Trim(lrCredPigJoyasDet!cKilataje) = 21 Then
'                lnSum21k = lnSum21k + lrCredPigJoyasDet!npesoneto
'            End If
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
'lsSM = "S/. "
'lsCadImp = ""
'
'lsDescLote = lsDescLote & Chr$(13)
'
'prPersonas.MoveFirst
'
'Dim sSql As String, RSTEMP As ADODB.Recordset, cn As COMConecta.DCOMConecta
'
' Set RSTEMP = New ADODB.Recordset
' Set cn = New COMConecta.DCOMConecta
'  sSql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
'  sSql = sSql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre,cPersNombre as cPersNombreLargo "
'  sSql = sSql & " from persona where cperscod='" & prPersonas!cPersCod & "'"
'  If RSTEMP.State = 1 Then RSTEMP.Close
'  cn.AbreConexion
'  Set RSTEMP = cn.CargaRecordSet(sSql)
'  If RSTEMP.State = 1 Then
'    If RSTEMP.RecordCount > 0 Then
'       lsApellido = RSTEMP!cPersApellido
'       lsNombre = RSTEMP!cPersNombre
'       lsNombreLargo = RSTEMP!cPersNombreLargo
'    End If
'    If RSTEMP.State = 1 Then RSTEMP.Close
'  End If
'  cn.CierraConexion
'  Set cn = Nothing
'
'    lsCodPers = prPersonas!cPersCod
'    lsNroDocIdentidad = IIf(IsNull(prPersonas!NroDNI), "", prPersonas!NroDNI)
'    lsNroDocTributario = IIf(IsNull(prPersonas!NroRUC), "", prPersonas!NroRUC)
'    lsDireccion = prPersonas!cPersDireccDomicilio
'    lszona = Mid(Trim(prPersonas!Zona), 1, 25)
'    lsCiudad = Mid(Trim(prPersonas!PROV), 1, 25) & " - " & Mid(Trim(prPersonas!DPTO), 1, 20)
'    lsTelefono = IIf(IsNull(prPersonas!cPersTelefono), "", prPersonas!cPersTelefono)
'
'    lnTicTEA = Round((((1 + lnTasInt / 100) ^ 12 - 1) * 100), 2)
'
'    lsCadImp = ""
'
'    lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
'    lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
'
'
'    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 8)
'    lsCadImp = lsCadImp & Space(20) & lsCodPers & Space(21) & Format(psFecContrato, "dd/MM/yyyy") & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(54) & Format(psFecVencimiento, "dd/MM/yyyy") & Space(14)
'
'    lsCadImp = lsCadImp & oFunI.gPrnBoldON & psCodCta & oFunI.gPrnBoldOFF & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNombreLargo) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsNroDocIdentidad) & " " & ofun.ImpreCarEsp(lsNroDocTributario) & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreFormat(ofun.ImpreCarEsp(lsDireccion), 50, 0) & Space(7) & ofun.ImpreCarEsp(Str(pnPlazo)) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lszona) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsCiudad) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(26) & ofun.ImpreCarEsp(lsTelefono) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnPiezas, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoPrestado, 12, 2, True) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroBruto, 12, 2, True) & Space(34) & ofun.ImpreFormat(lnTicTEA, 12, 2) & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(30) & ofun.ImpreFormat(pnOroNeto, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnInteresDif, 12, 2, True) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & Space(26) & lsSM & ofun.ImpreFormat(pnValorTasacion, 12, 2, True) & Space(30) & lsSM & ofun.ImpreFormat(pnMontoRec, 12, 2, True) & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    'Titulo y Monto Prestamo
'    cTempoTexto75 = ""
'    'cTempoTexto75 = "Item" & Space(3) & "Descripción" & Space(22) & "P. Neto"    '& Space(21) &oFun.ImpreFormat(pnInteresDif, 9, 2, True)
'
'    'Primer Item y el Prestamo
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(0), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnMontoPrestado, 18, 2, True) & oFunI.gPrnSaltoLinea
'
'    'Segundo Item y interes
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(1), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnInteresDif, 18, 2, True) & oFunI.gPrnSaltoLinea
'
'    'tercer Item y Costo de Tasacion
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(2), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoTasacion, 18, 2, True) & oFunI.gPrnSaltoLinea  '4
'
'    'Cuarto Item y Costo de Custodia
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(3), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnCostoCustodia, 18, 2, True) & oFunI.gPrnSaltoLinea  '5
'
'    'Quinto Item y el Impuesto
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(4), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea '& Space(22) & ofun.ImpreFormat(pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea  '6
'
'    'Item 6
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(5), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'
'    'Item 7
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(6), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'
'    'Item 8
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(7), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'
'    'Item 9
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(8), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & Space(22) & oFunI.gPrnSaltoLinea
'
'    'Item 10 y Neto a Recibir
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(9), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    'lsCadImp = lsCadImp & Space(18) & ofun.ImpreFormat(pnMontoRec - pnImpuesto, 18, 2, True) & oFunI.gPrnSaltoLinea
'
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(10), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(11), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(12), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(13), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(14), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(15), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(16), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(17), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(18), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'    cTempoTexto75 = ""
'    cTempoTexto75 = ofun.ImpreFormat(lmDetalle(19), 75)
'    lsCadImp = lsCadImp & Space(15) & cTempoTexto75 & oFunI.gPrnSaltoLinea
'
'    lsCadImp = lsCadImp & ofun.Linea(oFunI.gPrnSaltoLinea, "", 7)
'    lsCadImp = lsCadImp & Space(71) & oFunI.gPrnBoldON & psCodCta & oFunI.gPrnBoldOFF & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & Space(15) & "Venc." & Format(psFecVencimiento, "dd/MM/yyyy") & Space(10) & psCodUser & Space(33) & pnPlazo
'    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'
'    nImpreHojaResumenCredGaranMobiliariaOro = lsCadImp
'    Set oFunI = Nothing
'End Function

'CTI4 ERS0112020**********************
Public Function ImprimeBoletaAhorro(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional ByVal psCodCMAC As String, Optional ByVal pnMontoIFT As Double, Optional ByVal psCtaCodCred As String = "", Optional ByVal psNombreTitularEco As String = "", _
        Optional ByVal pbImpTMU As Boolean = False) As String

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
Dim sText As String

Dim oImpre As COMFunciones.FCOMImpresion ' CTI4 ERS0112020***


''  quitar este comentario

        Dim clsGen As COMDConstSistema.DCOMGeneral
        Set clsGen = New COMDConstSistema.DCOMGeneral
        
        Set oImpre = New COMFunciones.FCOMImpresion 'CTI4 ERS0112020***
 

ETIQ:

'On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)


 ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(64)

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(18)  ' cancela condensada
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(72)  ' desactiva negrita

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

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)

'Print #nFicSal, Chr$(10);
If psCodCMAC = "102" Then  ' CAJA LIMA
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & lsNegritaOff   'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))

                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
   
        If pbSaldoInt Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)  ''& Space(29 + sSep + sAux)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & Chr$(10)
        End If
        'Dim clsGen As COMDConstSistema.DCOMGeneral
        'Set clsGen = New COMDConstSistema.DCOMGeneral
        lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsmensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux

Else   '' Caja Trujillo
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Space(19 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Chr$(10)
        End If
        If Mid(psCodCta, 9, 1) = 1 Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Chr$(10)
            End If
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
            End If
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff  'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & lsNegritaOff   'Desactiva Negrita
            End If
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            End If
            lnNumLinCmac = 1
        End If
        
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Chr$(10)
        End If
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Chr$(10)
            End If
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Chr$(10)
            End If
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        If Not pbImpTMU Then
           ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr$(10)
        Else
          ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        End If
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                End If
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
                End If
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            End If
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        If pnMontoIFT > 0 Then
            sText = "ITF Efectivo"
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Space(-1 + 15) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Chr$(10)
            End If
        End If
        If pbSaldoInt Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
            End If
                If pbOpSaldoC Then
                    If Not pbImpTMU Then
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                    Else
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                    End If
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
 
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
                End If
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Chr$(10)
            End If
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Chr$(10)
        End If
        If psCodUsuRem = "" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Chr$(10)
            End If
        End If

        If Mid(Trim(psCtaCodCred), 6, 3) = "517" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp("Cliente : " & Left(Trim(psNombreTitularEco) & Space(40), 29)) & Space(14) & oImpre.ImpreCarEsp("Cliente : " & Trim(psNombreTitularEco)) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp("Cliente : " & Left(Trim(psNombreTitularEco) & Space(40), 29)) & Chr$(10)
            End If
        End If
        '*************************************************
        lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        If Not pbImpTMU Then
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsmensaje) & lsNegritaOff
        Else
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsmensaje) & lsNegritaOff
        End If
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux
        
End If
    Exit Function
End Function
Public Function nPrintReciboExtorCargoCta(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psCtaAhoExt As String, ByVal pnITFAho As Currency, ByVal psNombCliente As String, ByVal psOperacion As String, _
        ByVal pnMonto As Currency, ByVal pnSaldoCapAnt As Currency, ByVal pnMovExtornado As Long, _
        ByVal psCodUser As String, ByVal psVarPersCMAC As String, _
        Optional ByVal psmensaje As String = "", Optional ByVal psUsuOri As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False) As String
        
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim loConstSistema As New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
        
    Dim vEspacio As Integer
    vEspacio = 28
        
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    If pbImpTMU = False Then
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & Space(12) & Space(vEspacio + 4) _
                            & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Space(28 + vEspacio) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(5 + vEspacio) _
                                 & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & Space(vEspacio + 4) _
                            & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & Space(5 + vEspacio) _
                            & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & Space(9 + vEspacio) _
                            & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & Space(11 + vEspacio) _
                            & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & Space(11 + vEspacio) _
                            & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & Space(10 + vEspacio) _
                            & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Space(10 + vEspacio) _
                            & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Else
        lsCadImp = "" & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)

    End If
    nPrintReciboExtorCargoCta = lsCadImp
    Set oFunI = Nothing
End Function
'END CTI4

