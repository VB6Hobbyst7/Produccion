VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCredDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private vsCadImpresion As String
Private lsNegritaOn As String
Private lsNegritaOff  As String
Private lsSaltoLin As String
Private lsTab As String
Private EspHorxPag As Integer
Private EspVerxPag As Integer
Private nPuntPag As Integer
Private sCadCab As String

Private Function lnSaltoLinDoc() As String
    nPuntPag = nPuntPag + 1
    If nPuntPag > EspVerxPag Then
        nPuntPag = 0
        lnSaltoLinDoc = oImpresora.gPrnSaltoPagina & sCadCab
    Else
        lnSaltoLinDoc = oImpresora.gPrnSaltoLinea
    End If
    
End Function

Private Sub ImprimeCabeceraBoletaPago(ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, ByVal psNroTransac As String, ByVal psCtaCod As String, _
        ByVal psTipoCredito As String)
        psNomAge = "Sede Institucional"
        On Error GoTo ErrorImprimeCabeceraBoletaPago
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & ImpreFormat("CMACT - CREDITOS" & Space(5) & Mid(psTipoCredito, 1, 25), 53, 0) & ImpreFormat("CMACT - CREDITOS" & Space(5) & Mid(psTipoCredito, 1, 25), 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub

Public Sub ImprimeBoleta(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal psTipoCredito As String, ByVal pnCapPagado As Double, _
            ByVal pnIntCompPagado As Double, ByVal pnIntMorPagado As Double, ByVal pnGastoPagado As Double, _
            ByVal pnIntGraciaPagado As Double, ByVal pnIntSuspReprog As Double, ByVal pnSaldoCap As Double, _
            ByVal psProxFecPago As String, ByVal psCodUsu As String, ByVal sLpt As String)
            
Dim nFicSal As Integer

    On Error GoTo ErrorImprimeBoleta
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    
    Print #nFicSal, oImpresora.gPrnEspaLineaN;    'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, oImpresora.gPrnTamPagina22;  'Longitud de página a 22 líneas'
    Print #nFicSal, oImpresora.gPrnTamLetra10CPI;   'Tamaño 10 cpi
    Print #nFicSal, oImpresora.gPrnTpoLetraSansSerif;     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + oImpresora.gPrnCondensadaON  ' condensada
    'Print #nFicSal, oImpresora.gPrnCondensadaOFF  ' cancela condensada
    Print #nFicSal, oImpresora.gPrnBoldOFF  ' desactiva negrita
    
    Print #nFicSal, ""
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, psCuotasPagadas, _
            psFecha, psHora, psPersNombre, psNroTransac, psCtaCod, psTipoCredito)
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    Print #nFicSal, vsCadImpresion
    'Print #nFicSal, lsSaltoLin
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
    
    Exit Sub
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF
    lsSaltoLin = oImpresora.gPrnSaltoLinea
    EspHorxPag = 130
    EspVerxPag = 62
    nPuntPag = 0
    lsTab = Space(5)
End Sub

Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, Optional psTab As String = "", Optional pbCondensado As Boolean = True)
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If pbCondensado Then
        psCadImp = psCadImp & psTab & "Caja Municipal de Trujillo" & Space(80) & "Fecha :" & psFechaHora & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(61) & "USUARIO : " & psCodUsu & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2) & psTitulo & lnSaltoLinDoc
    Else
        psCadImp = psCadImp & psTab & "Caja Municipal de Trujillo" & Space(55) & "Fecha :" & psFechaHora & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2) & psTitulo & lnSaltoLinDoc
    End If
End Sub

Public Function ImprimirMetasAnalistas(ByVal psNomAge As String, _
    ByVal pdFechaHora As Date, ByVal psCodUsu As String)

Dim sCadImp As String
Dim RAnalista As ADODB.Recordset
Dim RDatos As ADODB.Recordset
Dim oCredGen As DCredGeneral
Dim oDCred As DCredito
Dim I As Integer
Dim oBase As DCredActualizaBD
Dim dFecSis As Date

    On Error GoTo ErrorImprimirMetasAnalistas
    Set oBase = New DCredActualizaBD
    dFecSis = oBase.dFechaHora
    Set oBase = Nothing
    dFecSis = CDate(Format(Format(pdFechaHora, "dd/mm/yyyy") & " " & Format(dFecSis, "hh:mm:ss"), "dd/mm/yyyy hh.mm:ss"))
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, Format(dFecSis, "dd/mm/yyyy hh:mm:ss"), psCodUsu, "LISTADO DE METAS POR ANALISTA")
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & ImpreFormat("Tipo", 12, 0) & ImpreFormat("Fec Inicio", 14, 0) & ImpreFormat("Fec Final", 22, 0) & _
        ImpreFormat("PYME", 12, 0) & ImpreFormat("Dscto Plan.", 16, 0) & ImpreFormat("Plazo Fijo", 17, 0) & ImpreFormat("CTS", 12, 0) _
        & ImpreFormat("USOS DIV.", 14, 0) & ImpreFormat("AGRICOLA", 14, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    
    Set oCredGen = New DCredGeneral
    Set RAnalista = oCredGen.CargaAnalistas
    Set oCredGen = Nothing
    Do While Not RAnalista.EOF
        Set oDCred = New DCredito
        Set RDatos = oDCred.RecuperaDatosMetasAnalistaReporte(RAnalista!cPersCod)
        Set oDCred = Nothing
        Do While Not RDatos.EOF
            If RDatos.Bookmark = 1 Then
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsNegritaOn
                sCadImp = sCadImp & "ANALISTA : " & PstaNombre(RAnalista!cPersNombre)
                sCadImp = sCadImp & lsNegritaOff & lnSaltoLinDoc
            End If
            sCadImp = sCadImp & ImpreFormat(Trim(RDatos!cTipoMeta), 7) & ImpreFormat(Format(RDatos!dInicial, "dd/mm/yyyy"), 12) _
                & ImpreFormat(Format(RDatos!dfinal, "dd/mm/yyyy"), 12)
            For I = 1 To 6
                sCadImp = sCadImp & ImpreFormat(RDatos!nMonto, 12)
                RDatos.MoveNext
                If RDatos.EOF Then
                    Exit Do
                End If
            Next I
            sCadImp = sCadImp & lnSaltoLinDoc
        Loop
        RDatos.Close
        Set RDatos = Nothing
        sCadImp = sCadImp & lnSaltoLinDoc
        RAnalista.MoveNext
    Loop
    RAnalista.Close
    Set RAnalista = Nothing
    ImprimirMetasAnalistas = sCadImp & oImpresora.gPrnSaltoPagina
    Exit Function

ErrorImprimirMetasAnalistas:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function ImprimeConsultaCredito(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
    ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psEstado As String, ByVal psLineaCred As String, ByVal psFteIngreso As String, _
    ByVal MatRelacCred As Variant, ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, _
    ByVal psApoderado As String, ByVal psTipoCuota As String, ByVal MatHitorial As Variant, ByVal MatDesembolsos As Variant, _
    ByVal MatCuotasPend As Variant, ByVal MatDeudaVenc As Variant, ByVal MatDeudaAFecha As Variant, ByVal psProtesto As String, _
    ByVal psCargoAutom As String, ByVal psRefinanciado As String) As String
Dim sCadImp As String
Dim I As Integer

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Resumen General de Credito", Space(5))
    sCadImp = sCadImp & lnSaltoLinDoc & Space(5) & lsNegritaOn & "Credito : " & psCtaCod & lsNegritaOff & Space(20) & "Tipo : " & psTipoCred & Space(20) & "Estado : " & psEstado & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CLIENTES RELACIONADOS" & Space(45) & "DATOS GENERALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(5) & String(60, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "Nombre" & Space(35) & "Relaciones" & Space(15) & "Linea de Credito      : " & psLineaCred & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(6) & "Fuente de Ingreso     : " & psFteIngreso & lnSaltoLinDoc
    For I = 0 To 4
        sCadImp = sCadImp & Space(5)
        If I <= UBound(MatRelacCred) - 1 Then
            sCadImp = sCadImp & ImpreFormat(MatRelacCred(I, 0), 40, 0) & ImpreFormat(MatRelacCred(I, 1), 20, 0)
        End If
        Select Case I
            Case 0
                sCadImp = sCadImp & Space(6) & "Analista Responsable  : " & psAnalista & lnSaltoLinDoc
            Case 1
                sCadImp = sCadImp & Space(6) & "Condicion del Credito : " & psCondicion & lnSaltoLinDoc
            Case 2
                sCadImp = sCadImp & Space(66) & "Destino del Credito   : " & psDestino & lnSaltoLinDoc
            Case 3
                sCadImp = sCadImp & Space(66) & "Tipo de Cuota         : " & psTipoCuota & lnSaltoLinDoc
        End Select
    Next I
    
    'Hitorial de Credito
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "HISTORIAL DE CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & Space(10) & ImpreFormat("FECHA", 12) & ImpreFormat("MONTO", 12) & ImpreFormat("Nro CUOTAS", 12) & ImpreFormat("PLAZO", 12) & ImpreFormat("CUOTA", 12) & ImpreFormat("P. GRACIA", 12) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatHitorial) - 1
        '                   Estado                                  Fecha                                   Monto                                   Nro Cuotas                              Plazo                                   Cuota                                   Periodo de Gracia
        sCadImp = sCadImp & Space(5) & ImpreFormat(MatHitorial(I, 0), 11, 0) & ImpreFormat(MatHitorial(I, 1), 16, 0) & ImpreFormat(MatHitorial(I, 2), 12, 0) & ImpreFormat(MatHitorial(I, 3), 16, 0) & ImpreFormat(MatHitorial(I, 4), 12, 0) & ImpreFormat(MatHitorial(I, 5), 18, 0) & ImpreFormat(MatHitorial(I, 6), 12, 0) & lnSaltoLinDoc
    Next I
    
    'DESEMBOLSOS DEL CREDITO
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "DESEMBOLSOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("GASTOS", 12, 0) & ImpreFormat("ESTADO", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatDesembolsos) - 1
        '                               FECHA                                       MONTO                                     GASTOS                                    ESTADO DEL DESEMBOLSO
        sCadImp = sCadImp & Space(5) & ImpreFormat(MatDesembolsos(I, 0), 19, 0) & ImpreFormat(MatDesembolsos(I, 1), 14, 0) & ImpreFormat(MatDesembolsos(I, 2), 12, 0) & ImpreFormat(MatDesembolsos(I, 3), 18, 0) & lnSaltoLinDoc
    Next I
    
    'CUOTAS PENDIENTES
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PENDIENTES" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & ImpreFormat("CUOTA", 10, 0) & ImpreFormat("FECHA", 17, 0) & ImpreFormat("CAPITAL", 14, 0) & ImpreFormat("INTERES", 19, 0) & ImpreFormat("MORA", 12, 0) & ImpreFormat("DIAS ATRASO", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatCuotasPend) - 1
        '                   Nro Cuota                                 FECHA DE VENCIMIENTO                      CAPITAL                                   INTERES                                   MORA                                      DIAS ATRASO
        sCadImp = sCadImp & Space(10) & ImpreFormat(MatCuotasPend(I, 0), 8, 0) & ImpreFormat(MatCuotasPend(I, 1), 18, 0) & ImpreFormat(MatCuotasPend(I, 2), 15, 0) & ImpreFormat(MatCuotasPend(I, 3), 18, 0) & ImpreFormat(MatCuotasPend(I, 4), 12, 0) & ImpreFormat(MatCuotasPend(I, 5), 12, 0) & lnSaltoLinDoc
    Next I
            
    'CUOTAS ATRASADAS Y PAGOS
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS ATRASADAS" & Space(45) & "PAGO A LA FECHA" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "CAPITAL      : " & ImpreFormat(CDbl(MatDeudaVenc(0)), 20) & Space(27) & "CAPITAL      : " & ImpreFormat(CDbl(MatDeudaAFecha(0)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "INTERES      : " & ImpreFormat(CDbl(MatDeudaVenc(1)), 20) & Space(27) & "INTERES      : " & ImpreFormat(CDbl(MatDeudaAFecha(1)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "MORA         : " & ImpreFormat(CDbl(MatDeudaVenc(2)), 20) & Space(27) & "MORA         : " & ImpreFormat(CDbl(MatDeudaAFecha(2)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "GASTOS       : " & ImpreFormat(CDbl(MatDeudaVenc(3)), 20) & Space(27) & "GASTOS       : " & ImpreFormat(CDbl(MatDeudaAFecha(3)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "TOTAL        : " & ImpreFormat(CDbl(MatDeudaVenc(4)), 20) & Space(27) & "TOTAL        : " & ImpreFormat(CDbl(MatDeudaAFecha(4)), 20) & lnSaltoLinDoc
    
    'sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, pscodusu, "Resumen General de Credito", Space(5))
    sCadImp = sCadImp & Space(5) & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "OTROS DATOS" & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(70, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "PROTESTO           : " & psProtesto & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "CARGO AUTOMATICO   : " & psCargoAutom & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "REFINANCIADO       : " & psRefinanciado & lnSaltoLinDoc
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    ImprimeConsultaCredito = sCadImp
End Function

Public Function ImprimeRegistroSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
    ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psTipoPers As String, _
    ByVal psCodCli As String, ByVal psNombreCli As String, ByVal psDocNat As String, _
    ByVal psDocJur As String, ByVal psDirecc As String, _
    ByVal psNroHijos As String, ByVal psEdad As String, ByVal psEstadoCiv As String, _
    ByVal psConyugue As String, ByVal psLECony As String, ByVal psRazonSoc As String, _
    ByVal psTipoFuente As String, ByVal psRazonSocDirecc As String, _
    ByVal psActividad As String, ByVal psIniAct As String, ByVal MatRelaCli As Variant, _
    ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, ByVal psLineaCred As String, _
    ByVal psFechaSol As String, ByVal MatDatosEstado As Variant, ByVal MatGarantias As Variant, _
    ByVal MatRelacInterna As Variant) As String
    
Dim sCadImp As String
Dim I As Integer

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "REGISTRO DE SOLICITUD DE CREDITO", Space(5), False)
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito : " & psCtaCod & Space(40) & psTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(40) & psTipoPers & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        :" & psCodCli & Space(5) & PstaNombre(psNombreCli) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   :" & psDocNat & Space(15) & "Doc. Juridico :" & psDocJur & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      :" & psDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Edad           :" & psEdad & " Años" & Space(30) & "Estado Civil : " & psEstadoCiv & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Conyugue       :" & ImpreFormat(psConyugue, 20, 2, 0) & " L.E:" & psLECony & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE INGRESO" & Space(30) & psTipoFuente & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Razon Social           :" & psRazonSoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion              :" & psRazonSocDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Actividad              :" & psActividad & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Inicio de Actividades  :" & psIniAct & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CODIGO", 10, 0) & ImpreFormat("NOMBRE", 30, 0) & ImpreFormat("RELACION", 12, 0) & ImpreFormat("DOC.IDENT.", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    For I = 0 To UBound(MatRelaCli) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatRelaCli(I, 0), 10, 0) & ImpreFormat(MatRelaCli(I, 1), 30, 0) & ImpreFormat(MatRelaCli(I, 2), 12, 0) & ImpreFormat(MatRelaCli(I, 3), 15, 0) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "DATOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Analista : " & psAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Destino : " & ImpreFormat(psDestino, 20, 0) & "Condicion : " & psCondicion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Linea Credito : " & ImpreFormat(psLineaCred, 20, 0) & "Fecha de Solicitud : " & psFechaSol & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & ImpreFormat("Estado", 20, 0) & ImpreFormat("Monto", 16, 0) & ImpreFormat("Cuota", 16, 0) & ImpreFormat("Cuotas", 15, 0) & ImpreFormat("Plazo", 10, 0) & ImpreFormat("Superavit", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    For I = 0 To 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatDatosEstado(I, 0), 15) & ImpreFormat(MatDatosEstado(I, 1), 15) & ImpreFormat(MatDatosEstado(I, 2), 15) & ImpreFormat(MatDatosEstado(I, 3), 15) & ImpreFormat(MatDatosEstado(I, 4), 7) & ImpreFormat(MatDatosEstado(I, 5), 15) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "GARANTIAS" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("Garantia", 30, 0) & ImpreFormat("Total", 20, 0) & ImpreFormat("Gravado", 20, 0) & lnSaltoLinDoc
    For I = 0 To UBound(MatGarantias) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatGarantias(I, 0), 30, 0) & ImpreFormat(MatGarantias(I, 1) & " " & MatGarantias(I, 2), 20, 0) & ImpreFormat(MatGarantias(I, 3) & " " & MatGarantias(I, 4), 20, 0) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(30) & lsNegritaOn & "RELACIONES INTERNAS (Creditos Vigentes)" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CLIENTE", 30, 0) & ImpreFormat("RELACION", 15, 0) & ImpreFormat("CREDITO", 20, 0) & ImpreFormat("ATRASO", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    If IsArray(MatRelacInterna) Then
        For I = 0 To UBound(MatRelacInterna) - 1
            sCadImp = sCadImp & lsTab & ImpreFormat(MatRelacInterna(I, 0), 30, 0) & ImpreFormat(MatRelacInterna(I, 1), 15, 0) & ImpreFormat(MatRelacInterna(I, 2), 20, 0) & ImpreFormat(MatRelacInterna(I, 3), 15, 0) & lnSaltoLinDoc
        Next I
    Else
        sCadImp = sCadImp & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    ImprimeRegistroSolicitud = sCadImp
End Function

Public Function ImprimeSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String) As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    On Error GoTo ErrorImprimeSolicitud
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosSolicitudCredito(psCtaCod)
    Set oDCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "SOLICITUD DE CREDITO", lsTab, False)
    If Not R.BOF And Not R.EOF Then
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & Space(30) & R!cTipoCred & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(5) & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Doc. Natural   : " & ImpreFormat(IIf(IsNull(R!DNI), "", R!DNI), 12, 0) & "Doc. Juridico    : " & ImpreFormat(IIf(IsNull(R!RUC), "", R!RUC), 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Nro Hijos      : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(IIf(IsNull(R!dPersNacCreac), CDate(psFecSis), R!dPersNacCreac)), "#0") & " Años " & Space(20) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE DE INGRESO" & Space(30) & IIf(IsNull(R!cFteInd), "DEPENDIENTE", "INDEPENDIENTE") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Razon Social           : " & R!cRazonSocial & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion              : " & R!cDireccRazon & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Actividad              : " & R!cCIIUdescripcion & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Inicio de Actividades  : " & Format(R!dPersNacCreac, "dd/mm/yyyy") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & lsNegritaOn & "DATOS PRINCIPALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Analista Responsable  : " & R!cAnalista & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Destino del Credito   : " & R!cDestino & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Condicion del Credito : " & R!cCondicion & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Fecha de Asignacion   : " & Format(R!dFecAsig, "dd/mm/yyyy") & lnSaltoLinDoc
        If Not IsNull(R!cNomInst) Then
            sCadImp = sCadImp & lsTab & "Institucion           : " & R!cNomInst & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "Codigo Modular        : " & IIf(IsNull(R!cCodModular), "", R!cCodModular) & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Monto Solicitado      : " & Format(R!nMontoSol, "#0.00") & Space(15) & R!cMonedaSol & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cuotas Solicitadas    : " & Trim(Str(R!nCuotasSol)) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Plazo Solicitado      : " & Trim(Str(R!nPlazoSol)) & " Dia(s)" & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    End If
    R.Close
    Set R = Nothing
    ImprimeSolicitud = sCadImp
    Exit Function

ErrorImprimeSolicitud:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ImprimeResumenComite(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String) As String
Dim sSql As String
Dim oDCredDoc As DCredDoc
Dim oDCred As DCredito
Dim oDGarantia As DGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String
    Set oDCredDoc = New DCredDoc
    Set R = oDCredDoc.RecuperaDatosResumenComite(psCtaCod)
    Set oDCredDoc = Nothing
    Set oDCred = New DCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing
    Set oDGarantia = New DGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE", lsTab, False)
    sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & Space(30) & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(2) & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   : " & R!DNI & Space(30) & "Doc. Juridico : " & R!RUC & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
    If IsNull(R!dPersNacCreac) Then
        sCadImp = sCadImp & lsTab & "Edad           : "
    Else
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(R!dPersNacCreac), "#0")
    End If
    sCadImp = sCadImp & Space(30) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Nro Hijos : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat(" CODIGO ", 15, 0) & ImpreFormat("NOMBRE", 40, 0) & ImpreFormat("RELACION", 20, 0) & ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & ImpreFormat(RRelaCred!cPersCod, 15, 0) & ImpreFormat(RRelaCred!cPersNombre, 40, 0) & ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
        End If
        RRelaCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & Space(30) & "DATOS DEL CREDITO" & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & ImpreFormat("ESTADO", 15, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CUOTA", 15, 0) & ImpreFormat("CUOTAS", 15, 0) & ImpreFormat("PLAZO", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & ImpreFormat(REstadosCred!cEstado, 15, 0) & ImpreFormat(REstadosCred!nMonto, 15, 2) & ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & ImpreFormat(REstadosCred!nCuotas, 15, 0) & ImpreFormat(REstadosCred!nPlazo, 15, 0) & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & "DESTINO : " & R!cDestino & Space(25) & "LINEA DE CREDITO : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & "GARANTIA : " & ImpreFormat(RGarantCred!cTpoGarantia, 20, 0) & "Garant. Total: " & ImpreFormat(RGarantCred!nPorGravar, 15, 0) & "Garant. Cred. : " & ImpreFormat(RGarantCred!nGravado, 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & Space(35) & "RESOLUCION DE COMITE" & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & Space(15) & "MONTO" & Space(5) & "# CUOTAS " & Space(10) & "MONTO CUOTA" & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APROBADO : " & ".... ...................... .......................           ....................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Tipo de Contrato :      Fondo : ................  Garantias : ................................................" & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(10) & "____________________________      ______________________________    _______________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    
    R.Close
    Set R = Nothing
    RRelaCred.Close
    Set RRelaCred = Nothing
    REstadosCred.Close
    Set REstadosCred = Nothing
    RGarantCred.Close
    Set RGarantCred = Nothing
    
    ImprimeResumenComite = sCadImp
End Function

Public Function ImprimeComprobanteDesembolso(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String) As String
    
Dim oCredDoc As DCredDoc
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim R As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RDesemb As ADODB.Recordset
Dim sCadImp As String

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocDesembolso(psCtaCod)
    Set oCredDoc = Nothing
    Set oCred = New DCredito
    Set REstadosCred = oCred.RecuperaEstadosdelCredito(psCtaCod, True, True, True)
    Set oCred = Nothing
    Set oCalend = New Dcalendario
    
    Set oCalend = New Dcalendario
    Set RDesemb = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "COMPROBANTE DE DESEMBOLSO", lsTab, False)
    sCadImp = sCadImp & lsTab & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito             : " & ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Fecha de Aprobacion : " & ImpreFormat(Format(R!dFecApr, "dd/mm/yyyy"), 20, 0) & "Codigo de Cliente : " & R!cPersCod & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("ESTADO", 20, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CUOTAS", 15, 0) & ImpreFormat("PLAZO", 15, 0) & ImpreFormat("CUOTA", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & ImpreFormat(REstadosCred!cEstado, 20, 0) & ImpreFormat(REstadosCred!nCuotas, 15, 0) & ImpreFormat(REstadosCred!nPlazo, 15, 0) & ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    REstadosCred.Close
    Set REstadosCred = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APODERADO         : " & PstaNombre(R!cApoderado) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA            : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO  : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INTERES           : " & Format(R!nTasaInteres, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TIPO DE DESEMBOLSO: " & R!cTipoDesemb & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(25) & "DESEMBOLSOS EFECTUADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(15) & ImpreFormat("FECHA", 15, 0) & ImpreFormat("CUOTA", 15, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("GASTOS", 15, 0) & ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & Space(15) & ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & ImpreFormat(Trim(Str(RDesemb!nCuota)), 15, 0) & ImpreFormat(RDesemb!nCapital, 15, 2) & ImpreFormat(RDesemb!nGasto, 15, 2) & IIf(RDesemb!nColocCalendEstado = gColocCalendEstadoPendiente, "PENDIENTE", "CANCELADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "_______________________________    ________________________________    _______________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & PstaNombre(R!cPersNombre) & Space(5) & R!cAnalista & Space(5) & "CAJA" & lnSaltoLinDoc
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    ImprimeComprobanteDesembolso = sCadImp
    
End Function

Public Function ImprimePlandePagos(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String) As String

Dim R As ADODB.Recordset
Dim oCredDoc As DCredDoc
Dim oNCred As NCredito
Dim MatCalend As Variant
Dim sCadImp As String
Dim I As Integer
Dim nIntereses As Double

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
    Set oCredDoc = Nothing
    
    Set oNCred = New NCredito
    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod)
    Set oNCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", lsTab, False)
    sCadImp = sCadImp & lsTab & "DUPLICADO" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cPersNombre) & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Analista      : " & PstaNombre(R!cAnalista) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & ImpreFormat(R!cTipoCuota, 20, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Interes       : " & ImpreFormat(Format(R!nTasaInteres, "#0.00"), 20, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Moneda        : " & ImpreFormat(R!cMoneda, 20, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("Fecha", 15, 0) & ImpreFormat("No Cuota", 10, 0) & ImpreFormat("Cuota", 12, 0) & ImpreFormat("Capital", 12, 0) & ImpreFormat("Interes", 12, 0) & ImpreFormat("Int.Gracia", 12, 0) & ImpreFormat("Gastos", 8, 0) & ImpreFormat("Saldo Cap.", 12, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    nIntereses = 0
    For I = 0 To UBound(MatCalend) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(Format(CDate(MatCalend(I, 0)), "ddd, d mmm yyyy"), 15, 0) & ImpreFormat(CInt(MatCalend(I, 1)), 4, 0) & ImpreFormat(CDbl(MatCalend(I, 3)) _
                                                                                                                                                                + CDbl(MatCalend(I, 4)) _
                                                                                                                                                                + CDbl(MatCalend(I, 5)) _
                                                                                                                                                                + CDbl(MatCalend(I, 6)) _
                                                                                                                                                                + CDbl(MatCalend(I, 7)) _
                                                                                                                                                                + CDbl(MatCalend(I, 8)), 10, 2) & ImpreFormat(CDbl(MatCalend(I, 3)), 9, 2) & ImpreFormat(CDbl(MatCalend(I, 4)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 5)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 9)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 10)), 10, 2) & lnSaltoLinDoc
        nIntereses = nIntereses + CDbl(MatCalend(I, 4))
        nIntereses = CDbl(Format(nIntereses, "#0.00"))
    Next I
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Totales : " & ImpreFormat(R!nMonto, 32, 2) & ImpreFormat(nIntereses, 7, 2) & ImpreFormat(R!nMonto + nIntereses, 32, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & oImpresora.gPrnTpoLetraCurier & oImpresora.gPrnSuperIdxOn
    Set oNCred = New NCredito
    sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(oNCred.TasaIntPerDias(R!nTasaInteres, 360) * 100, "#0.00") & "%" & lnSaltoLinDoc
    Set oNCred = Nothing
    sCadImp = sCadImp & oImpresora.gPrnSuperIdxOFF
    sCadImp = sCadImp & oImpresora.gPrnTpoLetraSansSerif
    
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        sCadImp = sCadImp & lsTab & "NOTA : " & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Si no se ha Realizado su Descuento por Planilla" & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Acerquese a nuestras Agencias a Cancelar y " & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Evite el cobro de Mora" & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    ImprimePlandePagos = sCadImp
End Function

Public Function ImprimePagoLote(ByVal MatPagos As Variant, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
    ByVal psNomAge As String, ByVal psFecha As String, ByVal psCodUsu As String, ByVal psInstitucion As String, ByVal psMoneda As String) As String
    
Dim I As Integer
Dim sCadImp As String
Dim oNCredito As NCredito
Dim nMontoPag As Double
Dim nAcumTot As Double
Dim nAcumCapPag As Double
Dim nAcumIntComPag As Double
Dim nAcumIntMorPag As Double

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Reporte de Pagos de Prestamos Personales", lsTab, True)
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & psInstitucion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & psMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 25, 0) & ImpreFormat("TITULAR", 30, 0) & ImpreFormat("CUOTA QUE PAGO", 25, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CAPITAL", 15, 0) & ImpreFormat("INTERES", 15, 0) & ImpreFormat("MORA", 10, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    
    Set oNCredito = New NCredito
    nAcumTot = 0
    nAcumCapPag = 0
    nAcumIntComPag = 0
    nAcumIntMorPag = 0
    For I = 0 To UBound(MatPagos) - 1
        nMontoPag = CDbl(MatPagos(I, 1))
        'nMontoPag = oNCredito.MatrizCapitalPagado(MatCalendDistrib) + oNCredito.MatrizGastoPag(MatCalendDistrib) + oNCredito.MatrizIntCompPagado(MatCalendDistrib) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib) + oNCredito.MatrizIntMorPagado(MatCalendDistrib)
        sCadImp = sCadImp & lsTab & ImpreFormat(MatPagos(I, 0), 21) & ImpreFormat(MatPagos(I, 3), 30) & ImpreFormat(oNCredito.MatrizCuotasPagadas(MatCalendDistrib(I)), 15) _
            & ImpreFormat(nMontoPag, 12) & ImpreFormat(oNCredito.MatrizCapitalPagado(MatCalendDistrib(I)), 14) & ImpreFormat(oNCredito.MatrizIntCompPagado(MatCalendDistrib(I)), 12) & ImpreFormat(oNCredito.MatrizIntMorPagado(MatCalendDistrib(I)), 9) & lnSaltoLinDoc
        nAcumTot = nAcumTot + nMontoPag
        nAcumCapPag = nAcumCapPag + oNCredito.MatrizCapitalPagado(MatCalendDistrib(I))
        nAcumIntComPag = nAcumIntComPag + oNCredito.MatrizIntCompPagado(MatCalendDistrib(I))
        nAcumIntMorPag = nAcumIntMorPag + oNCredito.MatrizIntMorPagado(MatCalendDistrib(I))
    Next I
    Set oNCredito = Nothing
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " PAGO TOTAL : " & Format(nAcumTot, "#0.00")
    sCadImp = sCadImp & lsTab & " CAPITAL PAGADO : " & Format(nAcumCapPag, "#0.00")
    sCadImp = sCadImp & lsTab & " INTERES PAGADO : " & Format(nAcumIntComPag, "#0.00")
    sCadImp = sCadImp & lsTab & " MORA PAGADA : " & Format(nAcumIntMorPag, "#0.00")
    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
    
    ImprimePagoLote = sCadImp
    
End Function

'*******************************************************************************************************
'Devuelve Un a Arreglo de Strings donde devuelve todos los documentos que se emiten en un desembolso
'Tanto en Abono a cuenta como en efectivo, el valor de 1 Indica papel boleta y el valor 2 indica papel carta
'*******************************************************************************************************
Public Function ImprimeDocumentosDesembolso(ByVal psCtaCod As String, ByVal pnMovNro As Long, ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pdFecSis As Date) As Variant
Dim MatDocs() As String
Dim MatTempDesem(1, 2) As String
Dim MatTempDesemGastos(1, 2) As String
Dim MatTempDesemCancel() As String
Dim MatTempComprob(1, 2) As String
Dim MatTempPlan(1, 2) As String
Dim nTotReg As Integer
Dim oDCredDoc As DCredDoc
Dim R As ADODB.Recordset
Dim sCadTmp As String
Dim AcumTmp As Double
Dim I As Integer
Dim nCapital As Double
Dim nGasto As Double
Dim nNumTran As Integer
    nTotReg = 0
    '*****************************************************************************
    '****   Recuperar Boletas de Desembolso
    '*****************************************************************************
    Set oDCredDoc = New DCredDoc
    Set R = oDCredDoc.RecuperaDatosBoletaDesembolso(pnMovNro, psCtaCod, psNomAge)
    sCadTmp = ""
    vsCadImpresion = ""
    nNumTran = 0
    Do While Not R.EOF
        If R.Bookmark = 1 Then
            nCapital = IIf(IsNull(R!nCapital), 0, R!nCapital)
        Else
            nGasto = IIf(IsNull(R!nGasto), 0, R!nGasto)
        End If
        nNumTran = nNumTran + R!nNumTran
        R.MoveNext
    Loop
    R.MoveFirst
    Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "")
    'Cuerpo de la Boleta
    
    sCadTmp = vsCadImpresion
    sCadTmp = sCadTmp & lsNegritaOn
    sCadTmp = sCadTmp & String(10, "-") & "COMPROBANTE DE DESEMBOLSO" & String(10, "-")
    sCadTmp = sCadTmp & Space(14) & String(10, "-") & "COMPROBANTE DE DESEMBOLSO" & String(10, "-") & lsSaltoLin
    
    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(13) & ImpreFormat(nCapital, 12)
    sCadTmp = sCadTmp & Space(16) & "Monto del Desembolso :" & Space(13) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & "Gastos:" & ImpreFormat(nGasto, 12)
    sCadTmp = sCadTmp & Space(16) & "Gastos :" & ImpreFormat(nGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(40, "-")
    sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
    sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(nCapital + nGasto, 12)
    sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(nCapital + nGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 12)
    sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
    
    R.Close
    
    MatTempDesem(0, 0) = "1"
    MatTempDesem(0, 1) = sCadTmp
    nTotReg = nTotReg + 1
    '*****************************************************************************
    '********** Recuperar Boletas de Gastos de Desembolso
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaGastosDesembolso(pnMovNro, psCtaCod, psNomAge)
    If Not R.BOF And Not R.EOF Then
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, "", psCtaCod, "")
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(10, "-")
        sCadTmp = sCadTmp & Space(14) & String(10, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(10, "-") & lsSaltoLin
    End If
    AcumTmp = 0
    Do While Not R.EOF
        sCadTmp = sCadTmp & Mid(R!cDescripcion, 1, 25) & Space(13) & ImpreFormat(R!nMonto, 12)
        sCadTmp = sCadTmp & Space(16) & Mid(R!cDescripcion, 1, 25) & Space(13) & ImpreFormat(R!nMonto, 12) & lsSaltoLin
        AcumTmp = AcumTmp + R!nMonto
        R.MoveNext
    Loop
    sCadTmp = sCadTmp & String(40, "-")
    sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
    sCadTmp = sCadTmp & "Total Gastos :" & Space(8) & ImpreFormat(AcumTmp, 12)
    sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(AcumTmp, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    R.MovePrevious
    sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
    R.Close
        
    MatTempDesemGastos(0, 0) = "1"
    MatTempDesemGastos(0, 1) = sCadTmp
    nTotReg = nTotReg + 1
    '*****************************************************************************
    '********** Recuperar Boletas de Cancelacion de Creditos
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaCancelCredDesembolso(pnMovNro, psCtaCod, psNomAge)
    ReDim MatTempDesemCancel(R.RecordCount, 2)
    Do While Not R.EOF
        sCadTmp = ""
        vsCadImpresion = ""
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), psCtaCod, "")
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-")
        sCadTmp = sCadTmp & Space(14) & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & "Capital :" & Space(13) & ImpreFormat(R!nCapital, 12)
        sCadTmp = sCadTmp & Space(16) & "Capital :" & Space(13) & ImpreFormat(R!nCapital, 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Interes :" & Space(13) & ImpreFormat(R!nInteres, 12)
        sCadTmp = sCadTmp & Space(16) & "Interes :" & Space(13) & ImpreFormat(R!nInteres, 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Gastos:" & ImpreFormat(R!nGasto, 12)
        sCadTmp = sCadTmp & Space(16) & "Gastos :" & ImpreFormat(R!nGasto, 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & String(40, "-")
        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(R!nCapital + R!nGasto + R!nInteres, 12)
        sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(R!nCapital + R!nGasto, 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsNegritaOff
        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(R!nSaldoCap, 12)
        sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(R!nSaldoCap, 12) & lsSaltoLin
        sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        nTotReg = nTotReg + 1
        R.MoveNext
    Loop
    R.Close
    
    'Recuperar Boletas de Abono a Cuenta
    'FALTA FUNCION DE RAUL.....................!
    
    '********************************************************************************
    'Recuperar Documento de Comprobante de Desembolso
    '********************************************************************************
    MatTempComprob(0, 0) = "2"
    MatTempComprob(0, 1) = ImprimeComprobanteDesembolso(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu)
    nTotReg = nTotReg + 1
    '********************************************************************************
    'Recuperar Documento de Plan de Pagos
    '********************************************************************************
    MatTempPlan(0, 0) = "2"
    MatTempPlan(0, 1) = ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu)
    nTotReg = nTotReg + 1
    
    ReDim MatDocs(nTotReg, 2)
    MatDocs(0, 0) = MatTempDesem(0, 0)
    MatDocs(0, 1) = MatTempDesem(0, 1)
    MatDocs(1, 0) = MatTempDesemGastos(0, 0)
    MatDocs(1, 1) = MatTempDesemGastos(0, 1)
    nTotReg = 2
    For I = 0 To UBound(MatTempDesemCancel) - 1
        MatDocs(I + 2, 0) = MatTempDesemCancel(R.Bookmark - 1, 0)
        MatDocs(I + 2, 1) = MatTempDesemCancel(R.Bookmark - 1, 1)
        nTotReg = nTotReg + 1
    Next I
    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
    MatDocs(nTotReg + 1, 0) = MatTempPlan(0, 0)
    MatDocs(nTotReg + 1, 1) = MatTempPlan(0, 1)
    
    ImprimeDocumentosDesembolso = MatDocs
    
End Function

Public Function ImprimeBoletaCmacLlam(ByVal psNomAge As String, ByVal pdFecSis As Date, _
    ByVal psNomCli As String, ByVal psNomCmac As String, ByVal pnMonto As Double, _
    ByVal psCodUser As String) As String

    ImprimeBoletaCmacLlam = ""
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat("CMACT", 20) & ImpreFormat(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"), 20) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat("CMACT", 20) & ImpreFormat(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"), 20) & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCli & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCli & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCmac & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCmac & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & "Monto :" & ImpreFormat(pnMonto, 40) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & "Monto :" & ImpreFormat(pnMonto, 40) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat(psCodUser, 40) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat(psCodUser, 40) & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & oImpresora.gPrnSaltoLinea
    
End Function

Public Function ImprimePagodeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String) As String
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim vsProducto As String
Dim vsFinanc As String
Dim vsOperacion As String

    Set oDCred = New DCredDoc
    Set R = oDCred.ReportePagoDeCreditos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INGRESOS POR PAGO DE CREDITOS DEL DIA " & Format(pdFecIni, "dd/mm/yyyy"), , True)
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE OPERACION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL OPERACION ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE FINANCIAMIENTO
            If Trim(R!COPECOD) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FUENTE FINANCIAMIENTO ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sFuentefinan = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        Exit Do
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 20) & _
                                    ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimePagodeCreditos = sCadImp
End Function


Public Function ImprimeDesembolsosEfectuados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProducts As Variant)
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosDesembolsados(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, pMatProducts)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS DESEMBOLSADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 12) & ImpreFormat("ANALISTA", 12) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE OPERACION
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL OPERACION ", 42) & ImpreFormat(R!nMontoApr, 20) & _
                                    ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 12) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12) & ImpreFormat(R!nTasa, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE FINANCIAMIENTO
            If Trim(R!COPECOD) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FUENTE FINANCIAMIENTO ", 42) & ImpreFormat(R!nMontoApr, 20) & _
                                    ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 12) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12) & ImpreFormat(R!nTasa, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sFuentefinan = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nMontoApr, 20) & _
                                    ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 12) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12) & ImpreFormat(R!nTasa, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL ", 42) & ImpreFormat(R!nMontoApr, 20) & _
                                    ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 12) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12) & ImpreFormat(R!nTasa, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL AGENCIA
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nMontoApr, 20) & _
                                        ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 12) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12) & ImpreFormat(R!nTasa, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 12) & ImpreFormat("ANALISTA", 12) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 12) & ImpreFormat("ANALISTA", 12) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 12) & ImpreFormat("ANALISTA", 12) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!COPECOD) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 12) & ImpreFormat("ANALISTA", 12) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 20) & _
                                    ImpreFormat(R!nMontoApr, 20) & _
                                    ImpreFormat(R!nDesembolso, 12) & ImpreFormat(R!nIntApr, 14) & ImpreFormat(R!Gasto, 12) & ImpreFormat(R!sAnalista, 12, 12) & R!nTasa & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeDesembolsosEfectuados = sCadImp
End Function


Public Function ImprimeSaldoCarteraVigente(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaSaldoCarteraVigente(pMatAgencias, pdFecIni, pnMoneda, psCodUser, pdFecSis, psNomAge)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SALDOS DE CARTERA VIGENTE AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE PLAZO
        If Trim(R!sFuentefinan) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PRODUCTO
            If Trim(R!sPlazo) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE AGENCIA
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL
                    If R!sAgencia = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        Exit Do
                    Else
                        nPuntPag = 0
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!sFuentefinan, 20) & ImpreFormat(R!nCasos, 15, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeSaldoCarteraVigente = sCadImp
End Function


Public Function ImprimeCreditosCancelados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosCanceladosTodos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS CANCELADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE Financiamiento
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE Producto
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'Final del Plazo
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIAS
                    If R!sPlazo = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 20) & _
                                    ImpreFormat(R!nDesembolso, 20) & _
                                    ImpreFormat(R!nCuotasApr, 12, 0) & ImpreFormat(R!nPlazoApr, 7, 0) & ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 15) & ImpreFormat(R!sAnalista, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosCancelados = sCadImp
End Function


Public Function ImprimeCreditosCanceladosConSaldo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosCanceladosConsaldo(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS CANCELADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE Financiamiento
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE Producto
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DEL PLAZO
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'Final de Agencias
                    If R!sPlazo = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 30) & _
                                    ImpreFormat(R!nSaldo, 10) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & ImpreFormat(R!cLineaCred, 12, 4) & ImpreFormat(R!sAnalista, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    ImprimeCreditosCanceladosConSaldo = sCadImp
End Function

Public Function ImprimeResumenSaldosCarteraXAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalista(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, pMatCond, pMatProd)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(R!nCar3Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 9)
                    sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(R!nCar3Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(R!cUser, 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(R!nCar3Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXAnalista = sCadImp
    
End Function

Public Function ImprimeMoraInstitucional(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, ByVal pbMoraAnt As Boolean) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraInstitucional(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, pMatCond, pMatProd, pbMoraAnt)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        For I = 0 To UBound(pMatCond) - 1
            If CInt(pMatCond(I)) = gColocCredCondNormal Then
                sCadImp = sCadImp & "/NORMAL"
            End If
            If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                sCadImp = sCadImp & "/PARALELO"
            End If
            If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                sCadImp = sCadImp & "/RECURRENTE"
            End If
        Next I
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
        If R!cRango = "RANGO1" Then
            sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
        Else
            If R!cRango = "RANGO2" Then
                sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
            Else
                If R!cRango = "RANGO3" Then
                    sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                Else
                    sCadImp = sCadImp & "MORA MAYOR A " & pnCar3I & " DIAS" & lnSaltoLinDoc
                End If
            End If
        End If
        sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        
        'Final de RANGO
        If R!cCtaCod = "TOTAL" Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL MORA ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & ImpreFormat(R!nDiasAtraso, 10, 0) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'Final de Agencias
            If R!cRango = "TOTAL" Then
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & ImpreFormat(R!nDiasAtraso, 10, 0) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & ImpreFormat(R!nDiasAtraso, 10, 0) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadImp = sCadImp & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadImp = sCadImp & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadImp = sCadImp & "/RECURRENTE"
                        End If
                    Next I
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                    If R!cRango = "RANGO1" Then
                        sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
                    Else
                        If R!cRango = "RANGO2" Then
                            sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
                        Else
                            If R!cRango = "RANGO3" Then
                                sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                            Else
                                sCadImp = sCadImp & "MORA MAYOR A " & pnCar3I & " DIAS" & lnSaltoLinDoc
                            End If
                        End If
                    End If
                    sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                If R!cRango = "RANGO1" Then
                    sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
                Else
                    If R!cRango = "RANGO2" Then
                        sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
                    Else
                        If R!cRango = "RANGO3" Then
                            sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                        Else
                            sCadImp = sCadImp & "MORA MAYOR A " & pnCar3I & " DIAS" & lnSaltoLinDoc
                        End If
                    End If
                End If
                sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30) & _
                                    ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(Trim(Str(R!nCuotaPend)) & "/" & Trim(Str(R!nCuotasApr)), 7, 7) & ImpreFormat(R!nDiasAtraso, 5, 0) & ImpreFormat(R!sAnalista, 12, 12) & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraInstitucional = sCadImp
End Function

Public Function ImprimeMoraXAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pMatAnalistas As Variant) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraporAnalista(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pMatAnalistas)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        For I = 0 To UBound(pMatCond) - 1
            If CInt(pMatCond(I)) = gColocCredCondNormal Then
                sCadImp = sCadImp & "/NORMAL"
            End If
            If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                sCadImp = sCadImp & "/PARALELO"
            End If
            If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                sCadImp = sCadImp & "/RECURRENTE"
            End If
        Next I
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("SALDO DE", 9) & ImpreFormat("CUOTA", 6, 0) & ImpreFormat("DIAS", 9, 1) & ImpreFormat("SALDO", 12) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("FECHA", 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("CUOTA", 6) & ImpreFormat(" MORA", 5) & ImpreFormat("ATRASO", 10, 0) & ImpreFormat("MORA", 12) & ImpreFormat("GASTO", 6, 0) & ImpreFormat("VENCIMIENTO", 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
    End If
    
    Do While Not R.EOF
        'FINAL DE ANALISTA
        If R!cCtaCod = "TOTAL" Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL ANALISTA ", 52) & ImpreFormat(R!nMontoDesemb, 10) & _
                    ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nSaldoCuota, 15) & ImpreFormat(R!nSaldoMora, 20) & ImpreFormat(R!nSaldoGasto, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE AGENCIA
            If R!cAnalista = "TOTAL" Then
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 52) & ImpreFormat(R!nMontoDesemb, 10) & _
                    ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nSaldoCuota, 15) & ImpreFormat(R!nSaldoMora, 20) & ImpreFormat(R!nSaldoGasto, 10) & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 52) & ImpreFormat(R!nMontoDesemb, 10) & _
                    ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nSaldoCuota, 15) & ImpreFormat(R!nSaldoMora, 20) & ImpreFormat(R!nSaldoGasto, 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadCab = lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadCab = sCadCab & "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadCab = sCadCab & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadCab = sCadCab & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadCab = sCadCab & "/RECURRENTE"
                        End If
                    Next I
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("SALDO DE", 9) & ImpreFormat("CUOTA", 6, 0) & ImpreFormat("DIAS", 9, 1) & ImpreFormat("SALDO", 12) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("FECHA", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("CUOTA", 6) & ImpreFormat(" MORA", 5) & ImpreFormat("ATRASO", 10, 0) & ImpreFormat("MORA", 12) & ImpreFormat("GASTO", 6, 0) & ImpreFormat("VENCIMIENTO", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & sCadCab
                End If
            Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadCab = "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadCab = sCadCab & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadCab = sCadCab & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadCab = sCadCab & "/RECURRENTE"
                        End If
                    Next I
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("SALDO DE", 9) & ImpreFormat("CUOTA", 6, 0) & ImpreFormat("DIAS", 9, 1) & ImpreFormat("SALDO", 12) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("FECHA", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("CUOTA", 6) & ImpreFormat(" MORA", 5) & ImpreFormat("ATRASO", 10, 0) & ImpreFormat("MORA", 12) & ImpreFormat("GASTO", 6, 0) & ImpreFormat("VENCIMIENTO", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & sCadCab
            End If
        End If
        
        If Not IsNull(R!nCuotaMora) Then
            sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30) & _
                        ImpreFormat(R!nMontoDesemb, 10) & _
                        ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuotasApr, 5, 0) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nCuotaMora, 5, 0) & ImpreFormat(R!nDiasAtraso, 5, 0) & ImpreFormat(R!nSaldoMora, 10) & ImpreFormat(R!nSaldoGasto, 10) & ImpreFormat(Format(R!dVenc, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraXAnalista = sCadImp
End Function

Public Function ImprimeCreditosProtestados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim sProducto As String
Dim nContProd As Long
Dim nContAge As Long

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosProtestados(pMatAgencias, pdFecIni, pnMoneda)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS PROTESTADOS AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & " Producto : " & Trim(R!cConsDescripcion)
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
        sProducto = R!cConsDescripcion
        nContProd = 0
        nContAge = 0
    End If
    Do While Not R.EOF
        'FINAL DE PRODUCTO
        If sProducto <> R!cConsDescripcion And sAgencia = R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = nContAge + nContProd
            nContProd = 0
            sProducto = R!cConsDescripcion
            sCadCab = " Producto : " & Trim(R!cConsDescripcion)
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        If sAgencia <> R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = nContAge + nContProd
            sProducto = R!cConsDescripcion
            
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContProd = 0
            nContAge = 0
            nPuntPag = 0
            sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & " Producto : " & Trim(R!cConsDescripcion)
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 25) & ImpreFormat(R!nSaldo, 10) & _
                        ImpreFormat(Mid(R!sFecha, 7, 2) & "/" & Mid(R!sFecha, 5, 2) & "/" & Mid(R!sFecha, 1, 4), 10) & _
                        ImpreFormat(R!nNroProxCuota, 4, 0) & ImpreFormat(Format(R!dVenc, "dd/mm/yyyy"), 10, 2) & ImpreFormat(R!nCuotaApr, 10) & ImpreFormat(R!nCuotaXAmort, 10) & ImpreFormat(R!nCuotaAmort, 10) & ImpreFormat(R!nDiasAtraso, 5, 0) & lnSaltoLinDoc
        nContProd = nContProd + 1
        R.MoveNext
    Loop
    
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge + nContProd & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    R.Close
    Set R = Nothing
    ImprimeCreditosProtestados = sCadImp
    
End Function

Public Function ImprimeCreditosRetirados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim nContAge As Long

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosRetirados(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, pMatCond, pMatProd)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS RETIRADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & "  AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("MONTO ", 11, 0) & ImpreFormat("CAUSA", 20, 0) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("ANALISTA", 12, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
        nContAge = 0
    End If
    Do While Not R.EOF
        If sAgencia <> R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = 0
            nPuntPag = 0
            sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("MONTO ", 11, 0) & ImpreFormat("CAUSA", 20, 0) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("ANALISTA", 12, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        nContAge = nContAge + 1
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 34) & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cConsDescripcion, 20) & ImpreFormat(Format(R!dPrdEstado, "dd/mm/yyyy"), 20) & ImpreFormat(R!cUser, 10)
        R.MoveNext
    Loop
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    R.Close
    Set R = Nothing
    ImprimeCreditosRetirados = sCadImp
End Function

Public Function ImprimeCartaMorosos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, _
    ByVal pnDiasAtrIni As Integer, ByVal pnDiasAtrFin As Integer, ByVal sFormato As String, ByVal pMatAnalistas As Variant) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sCadImpTmp As String
Dim sAgencia As String
Dim sCadFormato As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCartaCreditosMorosos(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pnDiasAtrIni, pnDiasAtrFin, pMatAnalistas)
    Set oDCred = Nothing
    sAgencia = ""
    sCadImp = ""
    sCadImpTmp = ""
    Do While Not R.EOF
        sCadFormato = sFormato
        sCadFormato = Replace(sCadFormato, "<<FECHA>>", Format(pdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cPersNombre), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<DIRECCION>>", Trim(R!cPersDireccDomicilio), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ZONA>>", Trim(R!Zona), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CREDITO>>", R!cCtaCod, , 3, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CUOTA>>", Trim(Str(R!nCuota)), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ANALISTA>>", PstaNombre(R!cAnalista), , 1, vbTextCompare)
        sCadImp = sCadImp & sCadFormato & oImpresora.gPrnSaltoPagina
        If R.Bookmark Mod 50 = 0 Then
            sCadImpTmp = sCadImpTmp & sCadImp
            sCadImp = ""
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImpTmp = sCadImpTmp & sCadImp
    ImprimeCartaMorosos = sCadImpTmp
    
End Function

Public Function ImprimeCartaInvitacionCreditoParalelo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, _
    ByVal pnNotaIni As Integer, ByVal pnNotaFin As Integer, ByVal sFormato As String, ByVal pMatAnalistas As Variant, _
    ByVal pnUltCuotas As Integer, ByVal pbPorc As Integer) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sCadImpTmp As String
Dim sAgencia As String
Dim sCadFormato As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCartaInvitacionCreditosParalelos(pMatAgencias, pnMoneda, pMatCond, pMatProd, pnNotaIni, pnNotaFin, pMatAnalistas, pnUltCuotas, pbPorc)
    Set oDCred = Nothing
    sAgencia = ""
    sCadImp = ""
    sCadImpTmp = ""
    Do While Not R.EOF
        sCadFormato = sFormato
        sCadFormato = Replace(sCadFormato, "<<FECHA>>", Format(pdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cPersNombre), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<DIRECCION>>", Trim(R!cPersDireccDomicilio), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ZONA>>", Trim(R!Zona), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CREDITO>>", R!cCtaCod, , 3, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CUOTA>>", Trim(Str(R!nCuotas)), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ANALISTA>>", PstaNombre(R!cAnalista), , 1, vbTextCompare)
        sCadImp = sCadImp & sCadFormato & oImpresora.gPrnSaltoPagina
        If R.Bookmark Mod 50 = 0 Then
            sCadImpTmp = sCadImpTmp & sCadImp
            sCadImp = ""
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImpTmp = sCadImpTmp & sCadImp
    ImprimeCartaInvitacionCreditoParalelo = sCadImpTmp
    
End Function


Public Function ImprimeCreditosXUbicacionGeo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal psUbicaciogeo As String) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosPorUbicacionGeo(pMatAgencias, pnMoneda, pMatCond, pMatProd, psUbicaciogeo)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES POR UBICACION GEOGRAFICA : " & Trim(Mid(psUbicaciogeo, 1, 35)), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("DIRECCION", 20, 0) & ImpreFormat("ZONA", 15, 0) & ImpreFormat("APROBADO", 10, 0) & ImpreFormat("SALDO CAP", 10, 0) & ImpreFormat("NOTA", 5, 0) & ImpreFormat("ANALISTA", 14, 0) & ImpreFormat("CUOTA", 14, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
    End If
    Do While Not R.EOF
        If sAgencia <> R!cAgeDescripcion Then
            nPuntPag = 0
            sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("DIRECCION", 20, 0) & ImpreFormat("ZONA", 15, 0) & ImpreFormat("APROBADO", 10, 0) & ImpreFormat("SALDO CAP", 10, 0) & ImpreFormat("NOTA", 5, 0) & ImpreFormat("ANALISTA", 14, 0) & ImpreFormat("CUOTA", 14, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 34) & ImpreFormat(R!cPersDireccDomicilio, 20) & ImpreFormat(R!Zona, 20) & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!nSaldo, 10) & ImpreFormat(IIf(IsNull(R!nColocNota), 0, R!nColocNota), 5, 0) & ImpreFormat(R!cUser, 5, 0) & ImpreFormat(R!nNroProxCuota & "/" & R!nCuotas, 5, 0) & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosXUbicacionGeo = sCadImp
End Function


Public Function ImprimeCreditosVigentes(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pnTipoCambio As Double, ByVal pnDiasAtrIni As Integer, ByVal pnDiasAtrFin As Integer) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosVigentes(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pnDiasAtrIni, pnDiasAtrFin, pnTipoCambio)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES AL: " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 20) & _
            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
            ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
            ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE Financiamiento
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 35) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 35) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sPlazo = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 35) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIAS
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 35) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 35) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 20) & _
                                ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                                ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                                ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                                ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                                ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                                ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 20) & _
                            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                            ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                            ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 20) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 20) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 10) & ImpreFormat(R!sTitular, 25) & _
                                    ImpreFormat(Mid(R!sInstitucion, 1, 22), 23) & _
                                    ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 12, 0) & ImpreFormat(R!sPlazoCuota, 5, 0) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) _
                                    & ImpreFormat(R!nCuotaMora, 8, 0) & ImpreFormat(R!nDiasAtraso, 8, 0) & ImpreFormat(R!nMontoCuota, 8) & ImpreFormat(Format(R!dFecUtlPago, "dd/mm/yyyy"), 12) _
                                    & ImpreFormat(R!IntDeveng, 6) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(IIf(IsNull(R!nCapVenc), 0, R!nCapVenc), 8) _
                                    & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & ImpreFormat(R!cEstadoDesc, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosVigentes = sCadImp

End Function


Public Function ImprimeCreditosXInstitucion(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pnTipoOrden As Integer, ByVal pnIncluyeMora As Integer) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosXInstitucion(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS PERSONALES POR INSTITUCION AL : " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 15) & ImpreFormat("COD. MODULAR", 20) & _
            ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
            ImpreFormat("CUOTA", 8) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                        ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 10) & ImpreFormat(R!sTitular, 25) & _
                                    ImpreFormat(R!cCodModular, 20) & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 12) & ImpreFormat(R!nMontoApr, 10) _
                                    & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!sCuotaPend, 6, 4) & ImpreFormat(R!nCuota, 6) & ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosXInstitucion = sCadImp

End Function

Public Function ImprimeMoraXInstitucion(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, pMatInstituciones As Variant) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraXInstituciones(pMatAgencias, pnMoneda, pdFecIni, pMatInstituciones)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS PERSONALES POR INSTITUCION AL : " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 19) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("CUOTA", 7) & ImpreFormat("DIAS", 12, 0) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat(" ", 39) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("ATRASO", 12, 0) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 29) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 19) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("CUOTA", 7) & ImpreFormat("DIAS", 12, 0) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat(" ", 39) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("ATRASO", 12, 0) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 19) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("CUOTA", 7) & ImpreFormat("DIAS", 12, 0) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat(" ", 39) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("ATRASO", 12, 0) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 10) & ImpreFormat(R!sTitular, 25) & _
                                    ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) _
                                    & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!sCuotaPend, 8, 4) & ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 0) & ImpreFormat(R!nSaldoMora, 10) & ImpreFormat(Format(R!dFecVencimiento, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraXInstitucion = sCadImp

End Function

Public Function ImprimeResumenSaldosCarteraXAnalistaConsumo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar4I As Integer) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalistaConsumo(pMatAgencias, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar4I, pMatCond, pMatProd)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar1I & "-" & pnCar1F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar2I & "-" & pnCar2F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > " & pnCar4I & "DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 9)
                    sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar1I & "-" & pnCar1F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar2I & "-" & pnCar2F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > " & pnCar4I & "DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(R!cUser, 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXAnalistaConsumo = sCadImp
    
End Function


Public Function ImprimeResumenSaldosCarteraXInstitucionConsumo(ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorInstitucionConsumo(pMatAgencias, pMatProd, pdFecSis)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR INSTITUCION AL " & Format(pdFecSis, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : SOLES " & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(32) & ImpreFormat("¦                  C A R T E R A    V I G E N T E               ¦          VENCIDOS > 30 DIAS     ¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("     INSTITUCION                     ¦           S O L E S           ¦        D O L A R E S          ¦     SOLES      ¦     DOLARES    ¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(32) & ImpreFormat("¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ CAP VENC.¦ NRO ¦ CAP VENC.¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!sInstitucion = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & oImpresora.gPrnSaltoPagina
                    sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : SOLES " & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(32) & ImpreFormat("¦                  C A R T E R A    V I G E N T E               ¦          VENCIDOS > 30 DIAS     ¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("     INSTITUCION                     ¦           S O L E S           ¦        D O L A R E S          ¦     SOLES      ¦     DOLARES    ¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(32) & ImpreFormat("¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ CAP VENC.¦ NRO ¦ CAP VENC.¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(PstaNombre(R!sInstitucion), 30)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXInstitucionConsumo = sCadImp
    
End Function

Public Function ImprimePagosConCheque(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, _
    ByVal pnTipoBusq As Integer, ByVal psNroDoc As String) As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim sInstitucion As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nTotalRep As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim nContTotalCasosRep As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoConCheque(pMatAgencias, pnMoneda, pMatCond, pMatProd, pdFecIni, pdFecFin, pnTipoBusq, psNroDoc)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS CON CHEQUE AL " & Format(pdFecSis, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
        sInstitucion = R!cInstitucion
    End If
    Do While Not R.EOF
        If sInstitucion <> R!cInstitucion Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            If sAgencia <> R!sAgencia Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 20)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                nContTotalCasosRep = nContTotalCasosRep + nContTotalCasos
                nTotalRep = nTotalRep + nTotal
                nTotal = 0
                nContTotalCasos = 0
                nContCasos = 0
                nSubTotal = 0
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                
                sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sAgencia = R!sAgencia
                sInstitucion = R!cInstitucion
            Else
                sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sAgencia = R!sAgencia
                sInstitucion = R!cInstitucion
            End If
                               
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(Trim(R!sTitular), 30) _
                & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 73)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nContTotalCasosRep = nContTotalCasosRep + nContTotalCasos
        nTotalRep = nTotalRep + nTotal
        nTotal = 0
        nContTotalCasos = 0
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasosRep, 5, 0) & ImpreFormat(nTotalRep, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        ImprimePagosConCheque = sCadImp
    End If
End Function

Public Function ImprimePagosDeOtraAgencia(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal psCodAge As String, ByVal psCodCmact As String) As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoDeOtrasAgencias(pnMoneda, pdFecIni, pdFecFin, psCodAge, psCodCmact)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS DE OTRAS AGENCIAS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("  Credito                       Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
    End If
    Do While Not R.EOF
        If sAgencia <> R!sAgencia Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
            sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("  Credito                       Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sAgencia = R!sAgencia
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 73)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    ImprimePagosDeOtraAgencia = sCadImp
End Function

Public Function ImprimePagosENOtrasAgencias(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal psCodAge As String, ByVal psCodCmact As String, _
    ByVal pMatProd As Variant, ByVal pMatCond As Variant) As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoENOtrasAgencias(pnMoneda, pdFecIni, pdFecFin, psCodAge, psCodCmact, pMatCond, pMatProd)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS EN OTRAS AGENCIAS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("  Credito                        Cliente                       Operacion                   Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
    End If
    Do While Not R.EOF
        If sAgencia <> R!sAgencia Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
            sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("  Credito                        Cliente                       Operacion                   Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sAgencia = R!sAgencia
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 30) & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 73)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    ImprimePagosENOtrasAgencias = sCadImp
End Function

Public Function ImprimeInteresesSuspenso(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant)
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaInteresesEnSuspenso(pMatAgencias, pdFecIni, pnMoneda, pMatProd)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INTERESES EN SUSPENSO DE CERDITOS REFINANCIADOS", , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE FUENTE DE FINANCIAMIENTO
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If Trim(R!sFuentefinan) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If Trim(R!sPlazo) = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIA
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30, 0) & _
                        ImpreFormat(Format(R!dRefinanciado, "dd/mm/yyyy"), 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nInteresComp + R!nInteresMor + R!nGastos, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeInteresesSuspenso = sCadImp
End Function

