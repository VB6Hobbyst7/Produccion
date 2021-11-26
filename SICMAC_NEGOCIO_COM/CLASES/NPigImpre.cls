VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPigImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsCadImpresion As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsSaltoLin As String

Public Function ImpreContratoPignoraticio(ByVal psCodCta As String, ByVal pbBuscaData As Boolean, Optional ByVal psCodPers As String, _
       Optional ByVal psNomPers As String, Optional ByVal psDocId As String, Optional ByVal psFonoPers As String, _
       Optional ByVal psDirPers As String, Optional ByVal pnPlazo As Integer, Optional ByVal psFecEmi As String, _
       Optional ByVal psFecVenc As String, Optional ByVal psTasador As String, Optional ByVal psMoneda As String, _
       Optional ByVal psAgencia As String, Optional ByVal prJoyas As Recordset, Optional ByVal pnPrestamo As Currency, _
       Optional ByVal pnComTasac As Currency, Optional ByVal pnNetoRecibir As Currency, Optional ByVal pnPiezas As Integer, _
       Optional ByVal pnPBruto As Double, Optional ByVal pnPNeto As Currency, Optional ByVal pnTasacion As Currency, _
       Optional ByVal pnComServicio As Currency, Optional ByVal pnTasaInt As Double, Optional ByVal pnNroDuplicado As Integer = 0, _
       Optional ByVal psFonoAge As String) As String

Dim lsCadImp As String
Dim lsDupli As String * 18
Dim lsTipo As String * 10, lsSubTipo As String * 11, lsMaterial As String * 11, lsObs As String * 35
Dim lnTasacion As Currency
Dim lnDuplicado As Integer
Dim lnFilas As Integer, I As Integer

If pbBuscaData = True Then

    Dim rsCredPig As ADODB.Recordset
    Dim rsCredPigPersona As ADODB.Recordset
    Dim oMuestraContrato As DPigContrato
    Dim oPigFunc As DPigFunciones
    
    Set oMuestraContrato = New DPigContrato
        Set rsCredPig = oMuestraContrato.dObtieneDatosContrato(psCodCta, 1)
        Set prJoyas = oMuestraContrato.dObtieneDetalleJoyas(psCodCta)
        Set rsCredPigPersona = oMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psCodCta)
    Set oMuestraContrato = Nothing

    If Not rsCredPigPersona.EOF Then
        psCodPers = rsCredPigPersona!cPersCod
        psNomPers = PstaNombre(rsCredPigPersona!cpersNombre)
        psDocId = IIf(IsNull(rsCredPigPersona!NroDNI), "", rsCredPigPersona!NroDNI)
        psFonoPers = IIf(IsNull(rsCredPigPersona!cPersTelefono), "", rsCredPigPersona!cPersTelefono)
        psDirPers = IIf(IsNull(rsCredPigPersona!cPersDireccDomicilio), "", rsCredPigPersona!cPersDireccDomicilio) + " " + IIf(IsNull(rsCredPigPersona!Zona), "", rsCredPigPersona!Zona)
    End If

    Set rsCredPigPersona = Nothing
    
    If Not rsCredPig.EOF Then
        pnPlazo = rsCredPig!nPlazo
        psFecEmi = rsCredPig!dVigencia
        psFecVenc = rsCredPig!dVenc
        psTasador = rsCredPig!nTasador
        psMoneda = "SOLES"
        pnPrestamo = rsCredPig!nMontoCol
        pnComTasac = rsCredPig!ComiTasacion
        pnNetoRecibir = pnPrestamo - pnComTasac
        pnPiezas = rsCredPig!npiezas
        pnPBruto = rsCredPig!nPBruto
        pnPNeto = rsCredPig!nPNeto
        pnTasacion = rsCredPig!nTasacion
        pnTasaInt = rsCredPig!nTasaInteres
        pnNroDuplicado = rsCredPig!nNroDuplic
        
    End If

    Set rsCredPig = Nothing
    
    Set oPigFunc = New DPigFunciones
    
    psFonoAge = oPigFunc.GetFonoAge(Mid(psCodPers, 4, 2))
    Set rsCredPig = oPigFunc.GetConceptoValor(gColPigConceptoCodComiServ)
    pnComServicio = rsCredPig!nValor
    Set rsCredPig = Nothing
    Set oPigFunc = Nothing
    
End If

If pnNroDuplicado > 0 Then
    lsDupli = "DUPLICADO No  " & CStr(pnNroDuplicado)
Else
    lsDupli = ""
End If
    
'========= Formacion de la Cadena de Impresion
lsCadImp = lsCadImp & Chr(10)
lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
lsCadImp = lsCadImp & Chr$(27) & Chr$(80)            'Tamaño  +...-: 80, 77, 103
lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
lsCadImp = lsCadImp & Space(52) & psCodCta & Chr(10)
lsCadImp = lsCadImp & Space(30) & ImpreCarEsp(lsDupli) & Chr(10)
lsCadImp = lsCadImp & Chr$(27) & Chr$(70) & Chr(10)   'Desactiva Negrita
lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
lsCadImp = lsCadImp & Chr$(27) & Chr$(77)            'Tamaño  +...-: 80, 77, 103

lsCadImp = lsCadImp & Space(60) & "PESO BRUTO : " & ImpreFormat(pnPBruto, 5, 2, True) & Chr(10) & Chr(10)
lsCadImp = lsCadImp & Space(12) & psDocId & Space(17) & psCodPers & Space(2) & ImpreCarEsp(psNomPers) & Chr(10)
lsCadImp = lsCadImp & Space(12) & ImpreCarEsp(psFonoPers) & Space(30) & ImpreCarEsp(psDirPers) & Chr(10) & Chr(10) & Chr(10)

lsCadImp = lsCadImp & Space(10) & ImpreFormat(pnPlazo, 2, 0, False) & Space(27) & ImpreFormat(pnTasaInt, 4, 2, False) & Space(14) & ImpreCarEsp(psAgencia) & Chr(10)
lsCadImp = lsCadImp & Space(10) & Format(psFecEmi, "dd/mm/yyyy") & Chr(10)
lsCadImp = lsCadImp & Space(10) & Format(psFecVenc, "dd/mm/yyyy") & Space(19) & ImpreFormat(pnComServicio, 4, 2, False) & Space(16) & psMoneda & Space(5) & psTasador & Chr(10)

lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10)

prJoyas.MoveFirst
Do While Not prJoyas.EOF
        
        If pbBuscaData > True Then
            lsTipo = prJoyas!Tipo
            lsSubTipo = prJoyas!SubTipo
            lsMaterial = prJoyas!Material
        Else
            lsTipo = prJoyas!Tipo
            If IsNull(prJoyas!SubTipo) Then
                lsSubTipo = ""
            Else
                lsSubTipo = prJoyas!SubTipo
            End If
            lsMaterial = prJoyas!Material
            
        End If
        lsObs = prJoyas!Observacion
        lnTasacion = CCur(IIf(prJoyas!Tasacion = "", 0, prJoyas!Tasacion)) + CCur(IIf(prJoyas!TasAdicion = "", 0, prJoyas!TasAdicion))
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
        lsCadImp = lsCadImp & Chr$(27) & Chr$(103)            'Tamaño  +...-: 80, 77, 103
        lsCadImp = lsCadImp & Space(2) & prJoyas!item & Space(2) & ImpreCarEsp(lsTipo) & Space(1) & ImpreCarEsp(lsMaterial) & Space(1)
        lsCadImp = lsCadImp & ImpreCarEsp(lsSubTipo) & Space(1) & ImpreFormat(prJoyas!PNeto, 3, 2, False) & Space(1) & "gr." & Space(3)
        lsCadImp = lsCadImp & ImpreCarEsp(lsObs) & Space(6) & ImpreFormat(lnTasacion, 6, 2, True) & Chr(10)
        
    lnFilas = lnFilas + 1
    prJoyas.MoveNext

Loop

For I = lnFilas To 23
    lsCadImp = lsCadImp & " " & Chr(10)
Next I

lsCadImp = lsCadImp & Chr$(27) & Chr$(107) & Chr$(2) 'Tipo letra : 0,1,2 - Roman,SansS,Courier
lsCadImp = lsCadImp & Chr$(27) & Chr$(77)            'Tamaño  +...-: 80, 77, 103

lsCadImp = lsCadImp & Space(1) & ImpreFormat(pnPrestamo, 6, 2, True) & Space(3) & ImpreFormat(pnComTasac, 4, 2, True) & Space(5)
lsCadImp = lsCadImp & ImpreFormat(pnNetoRecibir, 6, 2, True) & Space(6) & ImpreFormat(pnPiezas, 2, 0, False) & Space(5)
lsCadImp = lsCadImp & ImpreFormat(pnPBruto, 6, 2, True) & Space(5) & ImpreFormat(pnPNeto, 6, 2, True) & Space(5) & ImpreFormat(pnTasacion, 6, 2, True) & Chr(10)
lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10)

lsCadImp = lsCadImp & Chr$(27) & Chr$(69)   'Activa Negrita  75
lsCadImp = lsCadImp & Space(5) & "CONSULTA DE SALDOS, INF. GENERAL ATENCION AL CLIENTE TELEF.  " & psFonoAge
lsCadImp = lsCadImp & Chr$(27) & Chr$(70)

ImpreContratoPignoraticio = lsCadImp

Set prJoyas = Nothing

End Function

' *********** FALTA MODIFICAR ESTA FUNCION
Public Function ImpreReciboDuplicadoContrato(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
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
        lsCadImp = lsCadImp & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Space(13 + vEspacio) _
                            & " Tasa Efectiva Anual : " & Format(CalculaTasaEfectivaAnual(pnTasaInteresMensual / 100), "#0.00%") & Chr(10)
                            
        lsCadImp = lsCadImp & "" & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & "" & ImpreFormat(psMensaje, 40) & Chr(10)
        
        ImpreReciboDuplicadoContrato = lsCadImp
        
End Function

'***************************************************************************
' CAFF - 10/09/2002 Impresion de los Reportes para el Despacho de las Guias
'***************************************************************************
Public Function ImpreGuia(ByVal psNumGuia As String, ByVal pnTipoGuia As Integer, ByVal psNomCaja As String, ByVal psFecha As String) As String

Dim oPigImp As DPigImpresiones
Dim rsCab As Recordset
Dim Rs As Recordset
Dim lnPag As Integer
Dim lnLineas As Integer
Dim lnItem As Integer
Dim lnTPiezas As Integer
Dim lnPBruto As Currency
Dim lnPNeto As Currency
Dim lnTasacion As Currency
Dim lsCadImp

lsCadImp = ""
lnPag = 1
lnLineas = 0:   lnItem = 0

Set oPigImp = New DPigImpresiones

Set rsCab = oPigImp.GetDatosColocPigGuia(psNumGuia, pnTipoGuia)

If pnTipoGuia = 1 Then
    Set Rs = oPigImp.GetDatosColocPigGuiaDetLote(psNumGuia)
Else
    Set Rs = oPigImp.GetDatosColocPigGuiaDetPieza(psNumGuia)
End If

Set oPigImp = Nothing

If pnTipoGuia = 1 Then 'Impresion por Lote

    If Not Rs.EOF And Not Rs.BOF Then

        lsCadImp = ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
        lnLineas = 7

        Do While Not Rs.EOF

            lnItem = lnItem + 1
            lsCadImp = lsCadImp & ImpreFormat(lnItem, 3, 0, False) & Space(3) & Rs!cCtaCod & Space(3) & ImpreFormat(Rs!nTotPiezasCont, 2, 0, False)
            lsCadImp = lsCadImp & Space(5) & ImpreFormat(Rs!nPBruto, 4, 2, True) & Space(5) & ImpreFormat(Rs!nPNeto, 4, 2, True) & Space(5)
            lsCadImp = lsCadImp & ImpreFormat(Rs!Tasacion, 7, 2, True) & Space(5) & Rs!cUser & Chr(10)
            lnLineas = lnLineas + 1

            If lnLineas >= 58 Then
                lsCadImp = lsCadImp & Chr(12)
                lnPag = lnPag + 1
                lsCadImp = lsCadImp & ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
                lnLineas = 7
            End If

            lnTPiezas = lnTPiezas + Rs!nTotPiezasCont
            lnPBruto = lnPBruto + Rs!nPBruto
            lnPNeto = lnPNeto + Rs!nPNeto
            lnTasacion = lnTasacion + Rs!Tasacion

            Rs.MoveNext
        Loop

        If lnLineas >= 54 Then
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
            lnLineas = 7
        End If
            lsCadImp = lsCadImp & "TOTALES : " & Chr(10)
            lsCadImp = lsCadImp & Space(12) & "Cantidad :    " & ImpreFormat(lnItem, 4, 0, True) & " lote(s)" & Chr(10)
            lsCadImp = lsCadImp & Space(26) & ImpreFormat(lnTPiezas, 6, 0, True) & " pieza(s)" & Chr(10)
            lsCadImp = lsCadImp & Space(12) & "P.Bruto  :    " & ImpreFormat(lnPBruto, 12, 2, True) & Chr(10)
            lsCadImp = lsCadImp & Space(12) & "P.Neto   :    " & ImpreFormat(lnPNeto, 12, 2, True) & Chr(10)
            lsCadImp = lsCadImp & Space(12) & "Tasacion :    " & ImpreFormat(lnTasacion, 12, 2, True) & Chr(10)
    End If

Else    'Impresion por Pieza

    If Not Rs.EOF And Not Rs.BOF Then

        lsCadImp = ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
        lnLineas = 7

        Do While Not Rs.EOF

            lnItem = lnItem + 1
            lsCadImp = lsCadImp & ImpreFormat(lnItem, 3, 0, False) & Space(3) & Rs!cCtaCod & Space(3) & ImpreFormat(Rs!nItemPieza, 2, 0, False)
            lsCadImp = lsCadImp & Space(5) & ImpreFormat(Rs!nPesoBruto, 4, 2, True) & Space(5) & ImpreFormat(Rs!nPesoNeto, 4, 2, True) & Space(5)
            lsCadImp = lsCadImp & ImpreFormat(Rs!Tasacion, 7, 2, True) & Space(5) & Rs!cUser & Chr(10)
            lnLineas = lnLineas + 1

            If lnLineas >= 58 Then
                lsCadImp = lsCadImp & Chr(12)
                lnPag = lnPag + 1
                lsCadImp = lsCadImp & ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
                lnLineas = 7
            End If

            lnTPiezas = lnTPiezas + Rs!nItemPieza
            lnPBruto = lnPBruto + Rs!nPesoBruto
            lnPNeto = lnPNeto + Rs!nPesoNeto
            lnTasacion = lnTasacion + Rs!Tasacion

            Rs.MoveNext
        Loop

        If lnLineas >= 54 Then

            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & ImpreGuiaCab(psNumGuia, lnPag, rsCab!Motivo, rsCab!Origen, rsCab!Destino, pnTipoGuia, psNomCaja, psFecha)
            lnLineas = 7
        End If
        
            lsCadImp = lsCadImp & "TOTALES : " & Chr(10)
            lsCadImp = lsCadImp & Space(12) & "Cantidad :    " & ImpreFormat(lnItem, 4, 0, True) & "lote(s)" & Chr(10)
            lsCadImp = lsCadImp & Space(26) & ImpreFormat(lnTPiezas, 4, 0, False) & "pieza(s)" & Chr(10)

    End If

End If

Set rsCab = Nothing
Set Rs = Nothing

lsCadImp = lsCadImp & Chr(12)
lsCadImp = lsCadImp & ImpreGuiaRemision(psNumGuia, psFecha, pnTipoGuia)
ImpreGuia = lsCadImp

End Function

Public Function ImpreGuiaCab(ByVal psNumDoc As String, ByVal pnPag As Integer, ByVal psMotivo As String, ByVal psOrigen As String, _
                    ByVal psDestino As String, ByVal pnTipoGuia As Integer, ByVal psNomCaja As String, ByVal psFecha As String) As String

Dim sCadImp As String

sCadImp = sCadImp & psNomCaja & Space(30) & "DESPACHO DE GUIA Nro.  " & psNumDoc & Space(30) & "Pag. " & pnPag & Chr(10)
sCadImp = sCadImp & Space(40) & "MOTIVO :" & psMotivo & Space(30) & psFecha & Chr(10)
sCadImp = sCadImp & "ORIGEN  :  " & psOrigen & Space(15) & "DESTINO  :  " & psDestino & Chr(10)
sCadImp = sCadImp & String(120, "-") & Chr(10)

If pnTipoGuia = 1 Then
    sCadImp = sCadImp & "Sec" & Space(5) & "Lote" & Space(15) & "Pzs." & Space(5) & "P. Bruto" & Space(5) & "P. Neto" & Space(5) & "Tasacion" & Space(5) & "Tasador" & Chr(10)
Else
    sCadImp = sCadImp & "Sec" & Space(5) & "Pieza" & Space(15) & "Observacion" & Space(10) & "Material" & Space(5) & "P. Bruto" & Space(5) & "P. Neto" & Space(5) & "Tasacion" & Chr(10)
End If

sCadImp = sCadImp & String(120, "-") & Chr(10)

    ImpreGuiaCab = sCadImp
    
End Function

'***********************************************************************
' CAFF - 10/09/2002 Impresion de la Guia de Remision -  Remesas
'***********************************************************************
Public Function ImpreGuiaRemision(ByVal psNumDoc As String, ByVal psFecha As String, Optional ByVal pnTipoGuia As Integer) As String

Dim sCadImp As String
Dim oPigImp As DPigImpresiones
Dim rsCab As Recordset
Dim Rs As Recordset
Dim lsDescripcion As String * 30
Dim lnTPiezas As Long

    Set oPigImp = New DPigImpresiones
    
    Set rsCab = oPigImp.GetDatosColocPigGuiaRemCab(psNumDoc)
    Set Rs = oPigImp.GetDatosColocPigGuiaRemDet(psNumDoc, pnTipoGuia)
    
    Set oPigImp = Nothing
    
    sCadImp = sCadImp & Space(28) & rsCab!DirOrigen & Chr(10)
    sCadImp = sCadImp & Space(14) & rsCab!Destino & Space(50) & ArmaFecha(CDate(psFecha)) & Chr(10)
    sCadImp = sCadImp & Space(14) & rsCab!DirDestino & Chr(10) & Chr(10)
    sCadImp = sCadImp & Space(14) & rsCab!cPersCod & " - " & rsCab!cpersNombre & Chr(10) & Space(65) & "X"
    sCadImp = sCadImp & Space(14) & rsCab!cPersDireccDomicilio
    
    sCadImp = sCadImp & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    Do While Not Rs.EOF
        lsDescripcion = Rs!cConsDescripcion
        sCadImp = sCadImp & Space(15) & ImpreFormat(Rs!Piezas, 3, 0, False) & Space(8) & "Joyas de " & lsDescripcion & Space(10)
        sCadImp = sCadImp & ImpreFormat(Rs!Peso, 8, 2, True) & Chr(10)
        lnTPiezas = lnTPiezas + Rs!Piezas
        Rs.MoveNext
    Loop

    sCadImp = sCadImp & Space(13) & "_______" & Chr(10)
    sCadImp = sCadImp & Chr$(27) & Chr$(69)
    sCadImp = sCadImp & Space(13) & ImpreFormat(lnTPiezas, 6, 0, True) & Space(8) & "Joyas"
    sCadImp = sCadImp & Chr$(27) & Chr$(70)

Set rsCab = Nothing
Set Rs = Nothing

ImpreGuiaRemision = sCadImp

End Function
'***********************************************************************
' CAFF - 10/09/2002 Impresion de la Guia de Remision -  Remesas
'***********************************************************************
Public Function ImpreGuiaRecep(ByVal psNumDoc As String, ByVal pnTipoGuia As Integer, _
                                ByVal psFecha As String, ByVal psCodCMAC As String) As String

Dim sCadImp As String
Dim oPigImp As DPigImpresiones
Dim Rs As Recordset
Dim lsDescripcion As String * 30
Dim lnTPiezas As Long
Dim lsNumGuia As String * 20
Dim lsFecha As String * 20
Dim lsGuiaSunat As String * 20
Dim lsTipoGuia As String * 9

    Set oPigImp = New DPigImpresiones
    Set Rs = oPigImp.dObtieneGuiaRecep(psNumDoc)
    Set oPigImp = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        lsNumGuia = psNumDoc
        lsFecha = Format(psFecha, "dd/mm/yyyy")
        lsGuiaSunat = Rs!cSerDocSunat & "-" & Rs!cNumDocSunat
        If pnTipoGuia = 1 Then
            lsTipoGuia = " lote(s) "
        Else
            lsTipoGuia = " pieza(s)"
        End If
        
        sCadImp = sCadImp & psCodCMAC & Space(30) & "CUSTODIA DE VALORES" & Space(30) & psFecha & Chr(10)
        sCadImp = sCadImp & Space(40) & "CONSTANCIA DE RECEPCION" & Chr(10)
        sCadImp = sCadImp & String(90, "-") & Chr(10)
        sCadImp = sCadImp & "GUIA       : " & lsNumGuia & Space(10) & "MOTIVO   : " & ImpreCarEsp(Rs!Motivo) & Chr(10)
        sCadImp = sCadImp & "FECHA      : " & lsFecha & Space(10) & "ORIGEN   : " & ImpreCarEsp(Rs!Origen) & Chr(10)
        sCadImp = sCadImp & "GUIA SUNAT : " & lsGuiaSunat & Space(10) & "DESTINO  : " & ImpreCarEsp(Rs!Destino) & Chr(10)
        sCadImp = sCadImp & String(90, "-") & Chr(10) & Chr(10)
        
        sCadImp = sCadImp & "Cantidad   : " & ImpreFormat(Rs!nTotItem, 4, 0, True) & lsTipoGuia & Space(15) & "Tasacion  : " & ImpreFormat(Rs!nTotTasacion, 12, 2, True) & Chr(10)
        sCadImp = sCadImp & "Peso Bruto : " & ImpreFormat(Rs!nPesoBruto, 7, 2, True) & "gr(s)." & Chr(10)
        sCadImp = sCadImp & "Peso Neto  : " & ImpreFormat(Rs!nPesoNeto, 7, 2, True) & "gr(s)." & Chr(10)
        
                
    End If
    
Set Rs = Nothing

ImpreGuiaRecep = sCadImp

End Function

'**** CAFF
Public Function ImpreReciboCustodia(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, pnTotal As Currency, ByVal psCodUser As String, _
        ByVal psMensaje As String) As String
    
Dim lsCadImp As String

    Dim vEspacio As Integer
    vEspacio = 10
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
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & Space(5) & "------ COBRO DE CUSTODIA ------" & Space(20) & Space(vEspacio) _
                            & " ------ COBRO DE CUSTODIA ------" & Chr(10)
        lsCadImp = lsCadImp & " Custodia       : " & ImpreFormat(pnTotal, 22) & Space(vEspacio) _
                            & " Custodia       : " & ImpreFormat(pnTotal, 22) & Chr(10)
        lsCadImp = lsCadImp & " Total a Pagar  : " & ImpreFormat(pnTotal, 22) & Space(vEspacio) _
                            & " Total a Pagar  : " & ImpreFormat(pnTotal, 22) & Chr(10)
        lsCadImp = lsCadImp & " " & String(42, "-") & Space(vEspacio + 1) & String(42, "-") & Chr(10)
        
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(38) & Space(vEspacio) _
                            & " " & Format(psCodUser, "@@@@") & Chr(10)
                            
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 40) & Space(vEspacio + 1) _
                            & " " & ImpreFormat(psMensaje, 40) & Chr(10)
        
        ImpreReciboCustodia = lsCadImp
        
End Function

Public Function ImpreRepMartillero(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal pnNumRemate As Integer, _
                                    ByVal pnTipoProceso As Integer, ByVal psUbica As String, ByVal pdFecha As String) As String
Dim I As Integer
Dim ContLineas As Integer
Dim lnContPag As Integer
Dim lnNumLote As Integer
Dim lsImpre As String
Dim oDatos As DPigImpresiones
Dim Rs As Recordset
Dim rsDet As Recordset
Dim lnTRetasacion As Currency
Dim lnTValorDeuda As Currency
Dim lnTValorRemate As Currency
Dim lnTGramosOroPlata As Currency
Dim lnTGramosOtros As Currency
Dim lsTipo As String * 12
Dim lsSubTipo As String * 12
Dim lsMaterial As String * 12
Dim lsConserva As String * 8
Dim lsDescrip As String * 22
Dim lsDescripAdic As String * 22
Dim lsProceso As String * 20

    lnNumLote = 0
    lnContPag = 1
    lsImpre = ""
    lsImpre = Chr$(10)
    
    Select Case pnTipoProceso
    Case 1
        lsProceso = "PRIMERA SUBASTA"
    Case 2
        lsProceso = "SEGUNDA SUBASTA"
    Case 3
        lsProceso = "TERCERA SUBASTA"
    End Select
    
    lsImpre = lsImpre & Space(5) & psNomCmac & Space(139) & "Fecha : " & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10)
    lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "LISTADO PARA EL REMATE :  " & pnNumRemate & Space(85) & "Pag. " & lnContPag & Chr$(10)
    lsImpre = lsImpre & Space(80) & "OFICINA :  " & ImpreCarEsp(psUbica) & Chr$(10)
    lsImpre = lsImpre & Space(80) & lsProceso & Chr(10) & Chr(10)
    
    lsImpre = lsImpre & String(202, "-") & Chr$(10)
    lsImpre = lsImpre & ImpreFormat("Item", 4) & ImpreFormat("Tipo", 10) & ImpreFormat("SubTipo", 10) & ImpreFormat("Material", 14)
    lsImpre = lsImpre & ImpreFormat("Peso", 12) & ImpreFormat("Observacion", 25) & ImpreFormat("Observa Adic", 20) & ImpreFormat("Estado", 15)
    lsImpre = lsImpre & ImpreFormat("Peso", 10) & ImpreFormat("Valor", 10) & ImpreFormat("Valor Base", 18) & ImpreFormat("Valor", 18)
    lsImpre = lsImpre & ImpreFormat("VoBo", 15) & Chr(10)
    lsImpre = lsImpre & Space(46) & ImpreFormat("Bruto", 10) & Space(48) & ImpreFormat("Conservacion", 18) & ImpreFormat("Neto", 10)
    lsImpre = lsImpre & ImpreFormat("Gramo", 13) & ImpreFormat("Pieza", 15) & ImpreFormat("Venta", 12) & Chr(10)
    lsImpre = lsImpre & String(202, "-") & Chr$(10)
    
    ContLineas = 10
    
    Set oDatos = New DPigImpresiones
    Set Rs = oDatos.GetDatosCabMartillero(pnNumRemate, pnTipoProceso)
    
    If Not (Rs.EOF And Rs.BOF) Then
        Do While Not Rs.EOF
                
            lnNumLote = lnNumLote + 1
                
            lsImpre = lsImpre & Chr(10)
            lsImpre = lsImpre & "Lote Nro. " & lnNumLote & Space(4) & "Contrato : " & ImpreFormat(Rs!cCtaCod, 20) & "   Peso Neto : "
            lsImpre = lsImpre & ImpreFormat(Rs!PesoNeto, 4, 2, True) & "     Retasacion : " & ImpreFormat(Rs!Tasacion, 7, 2, True)
            lsImpre = lsImpre & "     Valor Deuda : " & ImpreFormat(Rs!ValorDeuda, 7, 2) & "     Valor Base Lote : "
            lsImpre = lsImpre & ImpreFormat(Rs!ValorBase, 7, 2) & Space(6) & "S/._____________" & Space(6) & "______________" & Chr(10)
            
            lnTRetasacion = lnTRetasacion + Rs!Tasacion
            lnTValorDeuda = lnTValorDeuda + Rs!ValorDeuda
            lnTValorRemate = lnTValorRemate + Rs!ValorBase
            
            ContLineas = ContLineas + 2
            
            If ContLineas >= 60 Then
            
                lsImpre = lsImpre & Chr(12)
                
                lnContPag = lnContPag + 1
                lsImpre = lsImpre & Space(5) & psNomCmac & Space(139) & "Fecha : " & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10)
                lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "LISTADO PARA EL REMATE :  " & pnNumRemate & Space(85) & "Pag. " & lnContPag & Chr$(10)
                lsImpre = lsImpre & Space(80) & "OFICINA :  " & ImpreCarEsp(psUbica) & Chr$(10)
                lsImpre = lsImpre & Space(80) & lsProceso & Chr(10) & Chr(10)
                
                lsImpre = lsImpre & String(202, "-") & Chr$(10)
                lsImpre = lsImpre & ImpreFormat("Item", 4) & ImpreFormat("Tipo", 10) & ImpreFormat("SubTipo", 10) & ImpreFormat("Material", 14)
                lsImpre = lsImpre & ImpreFormat("Peso", 12) & ImpreFormat("Observacion", 25) & ImpreFormat("Observa Adic", 20) & ImpreFormat("Estado", 15)
                lsImpre = lsImpre & ImpreFormat("Peso", 10) & ImpreFormat("Valor", 10) & ImpreFormat("Valor Base", 18) & ImpreFormat("Valor", 18)
                lsImpre = lsImpre & ImpreFormat("VoBo", 15) & Chr(10)
                lsImpre = lsImpre & Space(46) & ImpreFormat("Bruto", 10) & Space(48) & ImpreFormat("Conservacion", 18) & ImpreFormat("Neto", 10)
                lsImpre = lsImpre & ImpreFormat("Gramo", 13) & ImpreFormat("Pieza", 15) & ImpreFormat("Venta", 12) & Chr(10)
                lsImpre = lsImpre & String(202, "-") & Chr$(10)
            
                ContLineas = 11
            
            End If
            Set rsDet = oDatos.GetDatosDetMartillero(pnNumRemate, pnTipoProceso, Rs!cCtaCod)
            
            If Not (rsDet.EOF And rsDet.BOF) Then
                Do While Not rsDet.EOF
                        lsConserva = IIf(IsNull(rsDet!Conservacion), "", rsDet!Conservacion)
                        lsTipo = rsDet!TipoJoya
                        lsSubTipo = IIf(IsNull(rsDet!SubTipoJoya), "", rsDet!SubTipoJoya)
                        lsDescrip = IIf(IsNull(rsDet!cDescripcion), "", rsDet!cDescripcion)
                        lsDescripAdic = IIf(IsNull(rsDet!cDescripcionAdic), "", rsDet!cDescripcionAdic)
                        lsMaterial = IIf(IsNull(rsDet!Material), "", rsDet!Material)
                        lsImpre = lsImpre & ImpreFormat(rsDet!nItemPieza, 2, 0) & Space(3) & lsTipo
                        lsImpre = lsImpre & Space(3) & lsSubTipo & Space(3)
                        lsImpre = lsImpre & lsMaterial & Space(3) & ImpreFormat(rsDet!nPesoBruto, 5, 2, True)
                        lsImpre = lsImpre & Space(3) & lsDescrip & Space(3) & lsDescripAdic & Space(3)
                        lsImpre = lsImpre & lsConserva & Space(4)
                        lsImpre = lsImpre & ImpreFormat(rsDet!nPesoNeto, 8, 2, True) & ImpreFormat(rsDet!ValorGramo, 8, 2, True)
                        lsImpre = lsImpre & ImpreFormat(rsDet!nValorProceso, 11, 2, True)
                        lsImpre = lsImpre & Space(7) & "S/._____________" & Space(6) & "______________" & Chr(10)
                                    
                        If rsDet!nMaterial = 1 Or rsDet!nMaterial = 3 Or rsDet!nMaterial = 14 Or rsDet!nMaterial = 16 Or rsDet!nMaterial = 18 Or rsDet!nMaterial = 21 Or rsDet!nMaterial = 24 Then
                            lnTGramosOroPlata = lnTGramosOroPlata + rsDet!nPesoNeto
                        Else
                            lnTGramosOtros = lnTGramosOtros + rsDet!nPesoNeto
                        End If
                        
                        ContLineas = ContLineas + 1
                        
                        If ContLineas > 60 Then
                        
                            lsImpre = lsImpre & Chr(12)
                            
                            lnContPag = lnContPag + 1
                            
                            lsImpre = lsImpre & Space(5) & psNomCmac & Space(139) & "Fecha : " & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10)
                            lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "LISTADO PARA EL REMATE :  " & pnNumRemate & Space(85) & "Pag. " & lnContPag & Chr$(10)
                            lsImpre = lsImpre & Space(80) & "OFICINA :  " & ImpreCarEsp(psUbica) & Chr$(10)
                            lsImpre = lsImpre & Space(80) & lsProceso & Chr(10) & Chr(10)
                            
                            lsImpre = lsImpre & String(202, "-") & Chr$(10)
                            lsImpre = lsImpre & ImpreFormat("Item", 4) & ImpreFormat("Tipo", 10) & ImpreFormat("SubTipo", 10) & ImpreFormat("Material", 14)
                            lsImpre = lsImpre & ImpreFormat("Peso", 12) & ImpreFormat("Observacion", 25) & ImpreFormat("Observa Adic", 20) & ImpreFormat("Estado", 15)
                            lsImpre = lsImpre & ImpreFormat("Peso", 10) & ImpreFormat("Valor", 10) & ImpreFormat("Valor Base", 18) & ImpreFormat("Valor", 18)
                            lsImpre = lsImpre & ImpreFormat("VoBo", 15) & Chr(10)
                            lsImpre = lsImpre & Space(46) & ImpreFormat("Bruto", 10) & Space(48) & ImpreFormat("Conservacion", 18) & ImpreFormat("Neto", 10)
                            lsImpre = lsImpre & ImpreFormat("Gramo", 13) & ImpreFormat("Pieza", 15) & ImpreFormat("Venta", 12) & Chr(10)
                            lsImpre = lsImpre & String(202, "-") & Chr$(10)
                        
                            ContLineas = 10
                            
                           
                        End If
                        
                    rsDet.MoveNext
                Loop
            End If
            
            Set rsDet = Nothing
            Rs.MoveNext
        Loop
    End If
    
    If ContLineas > 60 Then
    
        lsImpre = lsImpre & Chr(12)
        
        lnContPag = lnContPag + 1
        lsImpre = lsImpre & Space(5) & psNomCmac & Space(139) & "Fecha : " & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10)
        lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "LISTADO PARA EL REMATE :  " & pnNumRemate & Space(85) & "Pag. " & lnContPag & Chr$(10)
        lsImpre = lsImpre & Space(80) & "OFICINA :  " & ImpreCarEsp(psUbica) & Chr$(10)
        lsImpre = lsImpre & Space(80) & lsProceso & Chr(10) & Chr(10)
        
        lsImpre = lsImpre & String(202, "-") & Chr$(10)
        lsImpre = lsImpre & ImpreFormat("Item", 4) & ImpreFormat("Tipo", 10) & ImpreFormat("SubTipo", 10) & ImpreFormat("Material", 14)
        lsImpre = lsImpre & ImpreFormat("Peso", 12) & ImpreFormat("Observacion", 25) & ImpreFormat("Observa Adic", 20) & ImpreFormat("Estado", 15)
        lsImpre = lsImpre & ImpreFormat("Peso", 10) & ImpreFormat("Valor", 10) & ImpreFormat("Valor Base", 18) & ImpreFormat("Valor", 18)
        lsImpre = lsImpre & ImpreFormat("VoBo", 15) & Chr(10)
        lsImpre = lsImpre & Space(46) & ImpreFormat("Bruto", 10) & Space(48) & ImpreFormat("Conservacion", 18) & ImpreFormat("Neto", 10)
        lsImpre = lsImpre & ImpreFormat("Gramo", 13) & ImpreFormat("Pieza", 15) & ImpreFormat("Venta", 12) & Chr(10)
        lsImpre = lsImpre & String(202, "-") & Chr$(10)
    
        ContLineas = 10
    End If
    
    lsImpre = lsImpre & Chr(10) & Chr(10)
    lsImpre = lsImpre & "      TOTALES : " & Chr(10)
    lsImpre = lsImpre & "      LOTES               :   " & ImpreFormat(lnNumLote, 9, 2) & Chr(10)
    lsImpre = lsImpre & "      GRAMOS ORO Y PLATA  :   " & ImpreFormat(lnTGramosOroPlata, 9, 2) & Chr(10)
    lsImpre = lsImpre & "      GRAMOS OTROS        :   " & ImpreFormat(lnTGramosOtros, 9, 2) & Chr(10)
    lsImpre = lsImpre & "      RETASACION          :   " & ImpreFormat(lnTRetasacion, 9, 2) & Chr(10)
    lsImpre = lsImpre & "      VALOR DEUDA         :   " & ImpreFormat(lnTValorDeuda, 9, 2) & Chr(10)
    lsImpre = lsImpre & "      VALOR REMATE        :   " & ImpreFormat(lnTValorRemate, 9, 2) & Chr(10)

    Set Rs = Nothing
    Set oDatos = Nothing
    ImpreRepMartillero = lsImpre

End Function

'MODIFICADO CRSF - 16/06/2003 - ADICION DE PARAMETRO OPCIONAL, LSPOLIZA
Public Function ImprePoliza(ByVal psNumPoliza As String, ByVal psCodPers As String, ByVal psNomPers As String, _
                            ByVal psDirPers As String, ByVal psFecha As String, pnValorVta As Currency, _
                            ByVal pnComision As Currency, ByVal pnIGV As Currency, ByVal pnTotal As Currency, _
                            ByVal psNumDI As String, ByVal prDetPoliza As Recordset, Optional ByVal lspoliza As String = "") As String

Dim lsImpre As String
Dim lsNomPers As String * 30
Dim lsNumDI As String * 10
Dim lsTipo As String * 15
Dim lsMaterial As String * 15
Dim lnLineas As Integer
Dim I As Integer

    lsNomPers = psNomPers
    lsImpre = ""
    lsImpre = lsImpre & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    lsImpre = lsImpre & Space(10) & Space(10) & Space(10) & Space(10) & Space(10) & "Nro - " & Mid(lspoliza, 4, 8) & Chr(10)
    lsImpre = lsImpre & Space(10) & lsNomPers & Space(10) & psCodPers & Chr(10)
    lsImpre = lsImpre & Space(10) & psDirPers & Chr(10)
    lsImpre = lsImpre & Space(10) & lsNumDI & Space(70) & psFecha & Chr(10)
    lsImpre = lsImpre & Chr(10) & Chr(10) & Chr(10)
    
    Do While Not prDetPoliza.EOF
        lsTipo = prDetPoliza!Tipo
        lsMaterial = prDetPoliza!Material
        lsImpre = lsImpre & prDetPoliza!Contrato & Space(7) & lsTipo & Space(7)
        lsImpre = lsImpre & lsMaterial & "  -  " & ImpreFormat(prDetPoliza!PesoNeto, 4, 2) & " gr."
        lsImpre = lsImpre & Space(35) & ImpreFormat(prDetPoliza!Importe, 5, 2, True) & Chr(10)
        lnLineas = lnLineas + 1
        prDetPoliza.MoveNext
    Loop

    Set prDetPoliza = Nothing
    For I = lnLineas To 20
        lsImpre = lsImpre & Chr(10)
    Next I

    lsImpre = lsImpre & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    lsImpre = lsImpre & Space(50) & ImpreFormat(pnValorVta, 7, 2, True) & Space(8) & ImpreFormat(pnComision, 7, 2, True)
    lsImpre = lsImpre & Space(10) & ImpreFormat(pnIGV, 7, 2, True) & Space(8) & "S/.   " & ImpreFormat(pnTotal, 7, 2, True) & Chr(10)
    
    ImprePoliza = lsImpre

End Function

Private Sub ImprimeCabBoletaPago(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psFecha As String, ByVal psHora As String, ByVal psPersNombre As String, ByVal psNroTransac As String, _
        ByVal psCtaCod As String, ByVal pnMovNro As Long, Optional ByVal psCodCMAC As String, Optional ByVal pcEmpresaRUC As String)
       
        On Error GoTo ErrorImprimeCabBoletaPago
        vsCadImpresion = ""
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 12 cpi
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        'Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(15)
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(70)  ' desactiva negrita
        
 If psCodCMAC = "102" Then
        vsCadImpresion = vsCadImpresion
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & psNomCmac & Space(8) & ImpreFormat("" & psFecha & Space(8) & "" & psHora, 38, 0) & "AG.:" & ImpreFormat(Mid(psNomAge, 9, 22), 25) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "R.U.C." & pcEmpresaRUC & CentrarCadena("CREDITO PIGNORATICIO", 28, 0) & ImpreFormat(psMoneda, 10, , False) & Space(4) & Trim(ImpreFormat(psNroTransac, 3, 2, False)) & lsNegritaOff & Chr$(27) + Chr$(69) & lsSaltoLin 'ImpreFormat(" - " & Space(17) & , 55, 0) &
        vsCadImpresion = vsCadImpresion & String(63, "-") & lsSaltoLin
        ''vsCadImpresion = vsCadImpresion & ImpreFormat(pnMovNro, 8, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69) & ImpreFormat("Cuenta       : " & psCtaCod, 33, 0) & Space(2) & ImpreFormat(Mid(psPersNombre, 1, 40), 35, 0) & Chr$(27) + Chr$(69) & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        
        '-centrarcadena(asd,)
        
'
'        vsCadImpresion = vsCadImpresion & psNomCmac & " - CREDITO PIGNORATICIO" & Space(8) & ImpreFormat(pnMovNro, 8, 0) & lsSaltoLin
'        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0) & lsSaltoLin
'        vsCadImpresion = vsCadImpresion & lsNegritaOff
'        vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(10) & "Hora : " & psHora, 53, 0) & lsSaltoLin
'        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
'        vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod, 53, 0) & lsSaltoLin
 Else
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & psNomCmac & " - CREDITO PIGNORATICIO" & Space(8) & ImpreFormat(pnMovNro, 8, 0) & lsSaltoLin
        '& Space(25) & psNomCmac & " - CREDITO PIGNORATICIO" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0) & lsSaltoLin
        '& Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(10) & "Hora : " & psHora, 53, 0) & lsSaltoLin
        '& Trim(ImpreFormat("Fecha : " & psFecha & Space(10) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
        '& Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod, 53, 0) & lsSaltoLin
        '& ImpreFormat("Credito : " & psCtaCod, 53, 0) & lsSaltoLin
  End If
        Exit Sub

ErrorImprimeCabBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabBoletaPago", Err.Description
    
End Sub

Private Sub Class_Initialize()
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
    lsSaltoLin = Chr$(10)
End Sub
                 
Public Sub ImpreReciboDesembolsoContrato(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnMontoPrestamo As Currency, _
        ByVal pnGastos As Currency, pnTotalPago As Currency, ByVal psCodUser As String, ByVal pnMovNro As Long, _
        ByVal sLpt As String, Optional ByVal psMensaje As String, Optional ByVal psCodCMAC As String, Optional ByVal pcEmpresaRUC As String)
            
Dim nFicSal As Integer

    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    

    Call ImprimeCabBoletaPago(psNomCmac, psNomAgencia, "SOLES", Format(psFecTransac, "dd/mm/yyyy"), Format(psFecTransac, "hh:mm:ss"), _
            psNombCliente, 1, psCodCta, pnMovNro, psCodCMAC, pcEmpresaRUC)
    
If psCodCMAC = "102" Then
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & "Transaccion  : COMPROBANTE DE DESEMBOLSO" & Space(7) & "Liq." & Space(2) & ImpreFormat(pnMovNro, 8, 0) & Chr$(27) & Chr$(70) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(63, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & " Monto Prestamo : " & ImpreFormat(pnMontoPrestamo, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Gastos         : " & ImpreFormat(pnGastos, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " Total a Pagar  : " & ImpreFormat(pnTotalPago, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " " & String(63, "-") & Chr(10)
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & Chr(10)
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & Chr(10)

Else
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(8, "-") & "COMPROBANTE DE DESEMBOLSO" & String(8, "-") & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(10) & String(8, "-") & "COMPROBANTE DE DESEMBOLSO" & String(8, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & " Monto Prestamo : " & ImpreFormat(pnMontoPrestamo, 22, 2, True) & Chr(10)
    '& Space(10) _                            & " Monto Prestamo : " & ImpreFormat(pnMontoPrestamo, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Gastos         : " & ImpreFormat(pnGastos, 22, 2, True) & Chr(10)
    '& Space(10) _   & " Gastos         : " & ImpreFormat(pnGastos, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " Total a Pagar  : " & ImpreFormat(pnTotalPago, 22, 2, True) & Chr(10)
    '& Space(10) _   & " Total a Pagar  : " & ImpreFormat(pnTotalPago, 22, 2, True) & Chr(10) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " " & String(42, "-") & Chr(10)
    '& Space(10) & " " & String(42, "-") & Chr(10)
    
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & Chr(10)
    '& Space(38) & Space(10) _  & " " & Format(psCodUser, "@@@@") & Chr(10)
                        
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & Chr(10)
    '& Space(10 + 1) _  & " " & ImpreFormat(psMensaje, 40) & Chr(10)

End If

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
       
End Sub

Public Sub ImpreReciboAmortizacion(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psFecTransac As String, _
         ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psFecNewVcto As String, _
         ByVal pnDiasAtraso As Integer, ByVal pnCapital As Currency, ByVal pnIntCompensatorio As Currency, _
         ByVal pnIntMoratorio As Currency, ByVal pnComServicio As Currency, ByVal pnDerRemate As Currency, _
         ByVal pnComVencido As Currency, ByVal pnMontoPagado As Currency, ByVal pnNewSaldoCap As Currency, _
         ByVal psCodUser As String, ByVal pnNroTransac As Long, ByVal pnMovNro As Long, ByVal pnNewPagoMin As Currency, _
         ByVal pnDiasIC As Integer, ByVal sLpt As String, Optional ByVal psMensaje As String = "", Optional ByVal psCodCMAC As String, Optional ByVal pcEmpresaRUC As String, Optional ByVal pnITF As Double = 0#)
        
Dim nFicSal As Integer

    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
   ' Print #nFicSal, Chr$(27) & Chr$(80);   'Tamaño 12 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    'Print #nFicSal, Chr$(27) + Chr$(15)

'

    Call ImprimeCabBoletaPago(psNomCmac, psNomAgencia, "SOLES", Format(psFecTransac, "dd/mm/yyyy"), Format(psFecTransac, "hh:mm:ss"), psNombCliente, pnNroTransac, _
                            psCodCta, pnMovNro, psCodCMAC, pcEmpresaRUC)
        
    vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & "Transaccion  : COMPROBANTE DE AMORTIZACION" & Space(7) & "Liq." & Space(2) & ImpreFormat(pnMovNro, 8, 0) & Chr$(27) & Chr$(70) & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & impreformat(,"Transaccion : COMPROBANTE DE AMORTIZACION", & Space(10) & ImpreFormat(pnMovNro, 8, 0) & lsSaltoLin
                                        '' " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(63, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & " Capital          : " & ImpreFormat(pnCapital, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " IntCom(" & ImpreFormat(pnDiasIC, 3, 0, True) & " dias) : " & ImpreFormat(pnIntCompensatorio, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Mora (" & ImpreFormat(pnDiasAtraso, 3, 0, True) & " dias)  : " & ImpreFormat(pnIntMoratorio, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Com. Servicio    : " & ImpreFormat(pnComServicio, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Der. Remate      : " & ImpreFormat(pnDerRemate, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Com. Vencido     : " & ImpreFormat(pnComVencido, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " I.T.F.           : " & ImpreFormat(pnITF, 19, 2, True) & Chr(10) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Total a Pagar    : " & ImpreFormat(pnMontoPagado + pnITF, 19, 2, True) & Chr(10) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
    vsCadImpresion = vsCadImpresion & " Nuevo Saldo      : " & ImpreFormat(pnNewSaldoCap, 19, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Vcto : " & Format(psFecNewVcto, "dd/mm/yyyy") & Space(5) & "Pag.Min :" & ImpreFormat(pnNewPagoMin, 7, , True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " " & String(63, "-") & Chr(10)
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & Space(38) & Chr(10)
                        
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & Chr(10)

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
            
End Sub

Public Sub ImpreReciboCancelacion(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psFecTransac As String, _
         ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnDiasAtraso As Integer, _
         ByVal pnCapital As Currency, ByVal pnIntCompensatorio As Currency, ByVal pnIntMoratorio As Currency, _
         ByVal pnComServicio As Currency, ByVal pnDerRemate As Currency, ByVal pnComVencido As Currency, _
         ByVal pnPenalidad As Currency, ByVal pnMontoPagado As Currency, ByVal psCodUser As String, _
         ByVal pnMovNro As Long, ByVal pnNroTransacc As Integer, ByVal sLpt As String, Optional ByVal psMensaje As String = "")

Dim nFicSal As Integer

    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

    Call ImprimeCabBoletaPago(psNomCmac, psNomAgencia, "SOLES", Format(psFecTransac, "dd/mm/yyyy"), Format(psFecTransac, "hh:mm:ss"), _
                psNombCliente, pnNroTransacc, psCodCta, pnMovNro)
        
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(7, "-") & "COMPROBANTE DE CANCELACION" & String(7, "-") & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(13) & String(7, "-") & "COMPROBANTE DE CANCELACION" & String(7, "-") & lsSaltoLin
        
    vsCadImpresion = vsCadImpresion & " Capital        : " & ImpreFormat(pnCapital, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Capital        : " & ImpreFormat(pnCapital, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Int. Compens.  : " & ImpreFormat(pnIntCompensatorio, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Int. Compens.  : " & ImpreFormat(pnIntCompensatorio, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Mora (" & ImpreFormat(pnDiasAtraso, 2, 0, True) & " dias) : " & ImpreFormat(pnIntMoratorio, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Mora (" & ImpreFormat(pnDiasAtraso, 2, 0, True) & " dias) : " & ImpreFormat(pnIntMoratorio, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Com. Servicio  : " & ImpreFormat(pnComServicio, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Com. Servicio  : " & ImpreFormat(pnComServicio, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Der. Remate    : " & ImpreFormat(pnDerRemate, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Der. Remate    : " & ImpreFormat(pnDerRemate, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Com. Vencido   : " & ImpreFormat(pnComVencido, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Com. Vencido   : " & ImpreFormat(pnComVencido, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Pen. Prto Pago : " & ImpreFormat(pnPenalidad, 21, 2, True) & lsSaltoLin
    '& Space(10) _  & " Pen. Prto Pago : " & ImpreFormat(pnPenalidad, 21, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " Total a Pagar  : " & ImpreFormat(pnMontoPagado, 21, 2, True) & lsSaltoLin
    '& Space(10) _   & " Total a Pagar  : " & ImpreFormat(pnMontoPagado, 21, 2, True) & Chr(10) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Desactiva Negrita
    vsCadImpresion = vsCadImpresion & " " & String(42, "-") & lsSaltoLin
    '& Space(9) & " " & String(42, "-") & Chr(10)
    
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & lsSaltoLin
    '& Space(38) & Space(10) _ & " " & Format(psCodUser, "@@@@") & Chr(10)
                        
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & lsSaltoLin
    '& Space(10 + 1) _ & " " & ImpreFormat(psMensaje, 40) & Chr(10)

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
                            
End Sub

' ****************************************************************************
'Modificacion de Impresion con adicion de Parametros
'CRSF - 25/09
' ****************************************************************************
Public Sub ImpreReciboDesembolsoRL(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal pnMontoPrestamo As Currency, _
        ByVal pnAmortizacion As Currency, ByVal pnComision As Currency, ByVal pnTotalPago As Currency, ByVal psCodUser As String, _
        ByVal pnMovNro As Long, ByVal pnNroTransac As Integer, ByVal sLpt As String, Optional ByVal psMensaje As String, _
        Optional pnValortasacion As Currency = 0)
            
Dim nFicSal As Integer

    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

    Call ImprimeCabBoletaPago(psNomCmac, psNomAgencia, "SOLES", Format(psFecTransac, "dd/mm/yyyy"), Format(psFecTransac, "hh:mm:ss"), psNombCliente, pnNroTransac, _
                            psCodCta, pnMovNro)
    
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(3, "-") & "COMPROBANTE DESEMBOLSO USO DE LINEA" & lsSaltoLin
    '& String(3, "-")    vsCadImpresion = vsCadImpresion & Space(10) & String(3, "-") & " COMPROBANTE DESEMBOLSO USO DE LINEA " & String(3, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & " Monto Prestamo : " & ImpreFormat(pnMontoPrestamo, 22, 2, True) & lsSaltoLin
    '& Space(10) _   & " Monto Prestamo : " & ImpreFormat(pnMontoPrestamo, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Amortizacion   : " & ImpreFormat(pnAmortizacion, 22, 2, True) & lsSaltoLin
    '& Space(10) _   & " Amortizacion   : " & ImpreFormat(pnAmortizacion, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Com. Servicio  : " & ImpreFormat(pnComision, 22, 2, True) & lsSaltoLin
    '& Space(10) _   & " Com. Servicio  : " & ImpreFormat(pnComision, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " Total a Pagar  : " & ImpreFormat(pnTotalPago, 22, 2, True) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & " Nueva Tasacion : " & ImpreFormat(pnValortasacion, 22, 2, True) & lsSaltoLin
    '& Space(10) _   & " Total a Pagar  : " & ImpreFormat(pnTotalPago, 22, 2, True) & Chr(10) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " " & String(42, "-") & lsSaltoLin
         '& Space(10) & " " & String(42, "-") & Chr(10)
    
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & lsSaltoLin
    '& Space(38) & Space(10) _ & " " & Format(psCodUser, "@@@@") & Chr(10)
                        
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & lsSaltoLin
    '& Space(10 + 1) _  & " " & ImpreFormat(psMensaje, 40) & Chr(10)

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
       
End Sub

Public Sub ImprimeBoletaExtorno(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psDescrip As String, ByVal psCtaCod As String, _
            ByVal psPersNombre As String, ByVal psFecha As String, ByVal psNroTransac As String, ByVal pnMontoTransac As Currency, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal pnMovNro As Long, ByVal sLpt As String)
            
Dim nFicSal As Integer

    On Error GoTo ErrorImprimeBoleta
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    Print #nFicSal, ""
    Call ImprimeCabBoletaPago(psNomCmac, psNomAge, "SOLES", Format(psFecha, "dd/mm/yyyy"), Format(psFecha, "hh:mm:ss"), _
                    psPersNombre, psNroTransac, psCtaCod, pnMovNro)
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(7, "-") & psDescrip & String(8, "-") & lsSaltoLin
    '& Space(10)  vsCadImpresion = vsCadImpresion & String(7, "-") & psDescrip & String(7, "-") & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado  : " & Space(12) & ImpreFormat(pnMontoTransac, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(10) & "Total Pagado  : " & Space(12) & ImpreFormat(pnMontoTransac, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital : " & Space(12) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(10) & "Saldo Capital : " & Space(12) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & String(42, "-") & lsSaltoLin
    '& Space(10) & String(42, "-") & Chr(10)
    
    vsCadImpresion = vsCadImpresion & Format(psCodUsu, "@@@@") & lsSaltoLin
    '& Space(38) & Space(10) _  & Format(psCodUsu, "@@@@") & Chr(10)
    
    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    
    Close #nFicSal
    
    Exit Sub
    
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    
End Sub

Public Sub ImpreReciboCancelacionSobrante(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal Rs As Recordset, ByVal psCtaAhorro As String, ByVal psNombCliente As String, ByVal pnMontoTransac As Currency, _
        ByVal pnSobrante As Currency, pnIntereses As Currency, ByVal psCodUser As String, _
        ByVal sLpt As String, Optional ByVal psMensaje As String)
            
Dim nFicSal As Integer
Dim I As Integer

    I = 1
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(72)   'desactiva negrita

    Call ImprimeCabBoletaPagoSobrante(psNomCmac, psNomAgencia, "SOLES", Format(psFecTransac, "dd/mm/yyyy"), Format(psFecTransac, "hh:mm:ss"), psNombCliente, 1, _
                            psCtaAhorro)
    
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(6, "-") & "COMPROBANTE DE PAGO DE SOBRANTE" & String(6, "-")
    vsCadImpresion = vsCadImpresion & Space(10) & String(6, "-") & "COMPROBANTE DE PAGO DE SOBRANTE" & String(6, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & Space(4) & "CONTRATO" & Space(23) & "SOBRANTE"
    vsCadImpresion = vsCadImpresion & Space(10) & Space(4) & "CONTRATO" & Space(23) & "SOBRANTE" & Space(4) & Chr(10)
    
    Rs.MoveFirst
    Do While Not Rs.EOF
        If I <= 2 Then
            vsCadImpresion = vsCadImpresion & Space(1) & Rs!Contrato & Space(6) & ImpreFormat(Rs!Sobrante, 15, 2, True) & Space(10) _
                                    & Space(1) & Rs!Contrato & Space(6) & ImpreFormat(Rs!Sobrante, 15, 2, True) & Chr(10)
                                
        End If
        I = I + 1
        Rs.MoveNext
    Loop
        
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & " Monto Sobrante : " & ImpreFormat(pnSobrante, 22, 2, True) & Space(10) _
                            & " Monto Sobrante : " & ImpreFormat(pnSobrante, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & " Interés        : " & ImpreFormat(pnIntereses, 22, 2, True) & Space(10) _
                            & " Interés        : " & ImpreFormat(pnIntereses, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(69)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " Total a Pagar  : " & ImpreFormat(pnMontoTransac, 22, 2, True) & Space(10) _
                        & " Total a Pagar  : " & ImpreFormat(pnMontoTransac, 22, 2, True) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(70)    'Activa Negrita
    vsCadImpresion = vsCadImpresion & " " & String(42, "-") & Space(10) & " " & String(42, "-") & Chr(10)
    
    vsCadImpresion = vsCadImpresion & " " & Format(psCodUser, "@@@@") & Space(38) & Space(10) _
                        & " " & Format(psCodUser, "@@@@") & Chr(10)
                        
    vsCadImpresion = vsCadImpresion & " " & ImpreFormat(psMensaje, 40) & Space(10 + 1) _
                        & " " & ImpreFormat(psMensaje, 40) & Chr(10)

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
       
End Sub

Private Sub ImprimeCabBoletaPagoSobrante(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psFecha As String, ByVal psHora As String, ByVal psPersNombre As String, ByVal psNroTransac As String, _
        ByVal psCtaCod As String)
       
        On Error GoTo ErrorImprimeCabBoletaPago
        vsCadImpresion = ""
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
        
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & psNomCmac & " - CREDITO PIGNORATICIO" & Space(25) & psNomCmac & " - CREDITO PIGNORATICIO" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(17) & Trim(ImpreFormat(psNroTransac, 2)), 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(10) & "Hora : " & psHora, 53, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(10) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat("Cuenta No : " & psCtaCod, 53, 0) & ImpreFormat("Cuenta No : " & psCtaCod, 53, 0) & lsSaltoLin
        Exit Sub

ErrorImprimeCabBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabBoletaPago", Err.Description
    
End Sub

Public Function ImpreAvisoSobrante(ByVal pnRemate As Long, ByVal pdFecha As String, psTextoCarta As String) As String

Dim sql As String
Dim Rs As Recordset
Dim oCon As DConecta
Dim lsCad As String
Dim lsCarta As String

Set oCon = New DConecta

    oCon.AbreConexion

    sql = "SELECT S.cCtaCod, PP.cPersCod, S.cCtaAbono, cPersNombre, nSobrante, cPersDireccDomicilio + ' ' + cUbiGeoDescripcion cDireccion, " _
        & "CodPostal = (SELECT cDesCodPostal FROM CODIGOPOSTAL WHERE SUBSTRING(cCodZon,1,7) = '3' + SUBSTRING(P.cPersDireccUbiGeo, 2,6)) " _
        & "FROM COLOCPIGSOBRANTE S INNER JOIN PRODUCTOPERSONA PP ON S.CCTACOD = PP.CCTACOD " _
        & "INNER JOIN PERSONA P ON PP.CPERSCOD = P.CPERSCOD " _
        & "INNER JOIN UBICACIONGEOGRAFICA UG ON P.CPERSDIRECCUBIGEO = UG.CUBIGEOCOD " _
        & "WHERE nRemate = " & pnRemate
        
    Set Rs = oCon.CargaRecordSet(sql)
            
    If Not (Rs.EOF And Rs.BOF) Then
            
        Do While Not Rs.EOF
            
             lsCarta = psTextoCarta
             lsCarta = Replace(lsCarta, "<<FECHA>>", Format(pdFecha, "dd/mm/yyyy"), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<NOMCLTE>>", PstaNombre(Rs!cpersNombre, False), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<DIRECCION>>", Rs!cDireccion, , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CODPOSTAL>>", Rs!CodPostal, , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CRED>>", Rs!cCtaCod, , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CODCLTE>>", Rs!cPersCod, , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<MONTO>>", ImpreFormat(Rs!nSobrante, 7, 2, True), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CUENTA>>", Rs!cCtaAbono, , 1, vbTextCompare)
             
             lsCad = lsCad & lsCarta & Chr(12)
         
             Rs.MoveNext
         Loop
                      
       ImpreAvisoSobrante = lsCad
               
    End If

End Function
























