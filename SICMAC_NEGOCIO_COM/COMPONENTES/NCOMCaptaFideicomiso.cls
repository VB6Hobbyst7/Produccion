VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaFideicomiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim vsCadImpresion As String

Public Sub CapPagoFideicomiso(ByVal psCodCta As String, ByVal pmPlanPagos As Variant, ByVal pnMontoPago As Currency, _
                              ByVal pnCuotasPag As Integer, ByVal pnTipoPago As Integer, ByVal psMovNro As String, _
                              ByVal pnTipoCalend As Integer, Optional ByVal pnCuotaPagPreJud As Integer)
            
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim lnMovNro As String
Dim i As Integer

Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobFideicomiso, "Pago de Fideicomiso", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For i = 1 To pnCuotasPag
       
        If pmPlanPagos(i, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(i, 1), pmPlanPagos(i, 2)
        End If
        
        If pmPlanPagos(i, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(i, 1), pmPlanPagos(i, 3)
        End If
        
        If pmPlanPagos(i, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(i, 1), pmPlanPagos(i, 4)
        End If
        
        If pmPlanPagos(i, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(i, 1), pmPlanPagos(i, 5)
        End If
        
        If pmPlanPagos(i, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(i, 1), pmPlanPagos(i, 6)
        End If
        
        If pmPlanPagos(i, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(i, 1), pmPlanPagos(i, 7)
        End If
        
        If pmPlanPagos(i, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(i, 1), pmPlanPagos(i, 8)
        End If
        
        If pmPlanPagos(i, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(i, 1), pmPlanPagos(i, 9)
        End If
        
        oGraba.dInsertFideicomisoMov psCodCta, pnTipoPago, pmPlanPagos(i, 1), pmPlanPagos(i, 2), pmPlanPagos(i, 3), pmPlanPagos(i, 6), _
                                     pmPlanPagos(i, 4), pmPlanPagos(i, 5), pmPlanPagos(i, 7), pmPlanPagos(i, 8), pmPlanPagos(i, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next i
    
    If pnTipoCalend = 0 Then
        oGraba.dUpdateFideicomiso psCodCta, pnCuotasPag, pnMontoPago
    Else
        oGraba.dUpdateFideicomiso psCodCta, pnCuotaPagPreJud, pnMontoPago
    End If
    
    oGraba.dCommitTrans

Set oGraba = Nothing

End Sub

'FONCODES
' CMCPL - 01/07
Public Sub CapPagoFONCODES(ByVal psCodCta As String, ByVal pmPlanPagos As Variant, ByVal pnMontoPago As Currency, _
                              ByVal pnCuotasPag As Integer, ByVal pnTipoPago As Integer, ByVal psMovNro As String, _
                              ByVal pnTipoCalend As Integer, Optional ByVal pnCuotaPagPreJud As Integer)
            
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim lnMovNro As String
Dim i As Integer

Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobFoncodes, "Pago de FONCODES", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For i = 1 To pnCuotasPag
       
        If pmPlanPagos(i, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(i, 1), pmPlanPagos(i, 2)
        End If
        
        If pmPlanPagos(i, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(i, 1), pmPlanPagos(i, 3)
        End If
        
        If pmPlanPagos(i, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(i, 1), pmPlanPagos(i, 4)
        End If
        
        If pmPlanPagos(i, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(i, 1), pmPlanPagos(i, 5)
        End If
        
        If pmPlanPagos(i, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(i, 1), pmPlanPagos(i, 6)
        End If
        
        If pmPlanPagos(i, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(i, 1), pmPlanPagos(i, 7)
        End If
        
        If pmPlanPagos(i, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(i, 1), pmPlanPagos(i, 8)
        End If
        
        If pmPlanPagos(i, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(i, 1), pmPlanPagos(i, 9)
        End If
        
        oGraba.dInsertFONCODESMov psCodCta, pnTipoPago, pmPlanPagos(i, 1), pmPlanPagos(i, 2), pmPlanPagos(i, 3), pmPlanPagos(i, 6), _
                                     pmPlanPagos(i, 4), pmPlanPagos(i, 5), pmPlanPagos(i, 7), pmPlanPagos(i, 8), pmPlanPagos(i, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next i
    
    If pnTipoCalend = 0 Then
        oGraba.dUpdateFONCODES psCodCta, pnCuotasPag, pnMontoPago
    Else
        oGraba.dUpdateFONCODES psCodCta, pnCuotaPagPreJud, pnMontoPago
    End If
    
    oGraba.dCommitTrans

Set oGraba = Nothing

End Sub
' PLAN BICI
' CMCPL - CRSF 01/07
Public Sub CapPagoPLANBICI(ByVal psCodCta As String, ByVal pmPlanPagos As Variant, ByVal pnMontoPago As Currency, _
                              ByVal pnCuotasPag As Integer, ByVal pnTipoPago As Integer, ByVal psMovNro As String, _
                              ByVal pnTipoCalend As Integer, Optional ByVal pnCuotaPagPreJud As Integer)
            
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim lnMovNro As String
Dim i As Integer

Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobPlanBici, "Pago de PLAN BICI", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For i = 1 To pnCuotasPag
       
        If pmPlanPagos(i, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(i, 1), pmPlanPagos(i, 2)
        End If
        
        If pmPlanPagos(i, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(i, 1), pmPlanPagos(i, 3)
        End If
        
        If pmPlanPagos(i, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(i, 1), pmPlanPagos(i, 4)
        End If
        
        If pmPlanPagos(i, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(i, 1), pmPlanPagos(i, 5)
        End If
        
        If pmPlanPagos(i, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(i, 1), pmPlanPagos(i, 6)
        End If
        
        If pmPlanPagos(i, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(i, 1), pmPlanPagos(i, 7)
        End If
        
        If pmPlanPagos(i, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(i, 1), pmPlanPagos(i, 8)
        End If
        
        If pmPlanPagos(i, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(i, 1), pmPlanPagos(i, 9)
        End If
        
        oGraba.dInsertPLANBICIMov psCodCta, pnTipoPago, pmPlanPagos(i, 1), pmPlanPagos(i, 2), pmPlanPagos(i, 3), pmPlanPagos(i, 6), _
                                     pmPlanPagos(i, 4), pmPlanPagos(i, 5), pmPlanPagos(i, 7), pmPlanPagos(i, 8), pmPlanPagos(i, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next i
    
    If pnTipoCalend = 0 Then
        oGraba.dUpdatePLANBICI psCodCta, pnCuotasPag, pnMontoPago
    Else
        oGraba.dUpdatePLANBICI psCodCta, pnCuotaPagPreJud, pnMontoPago
    End If
    
    oGraba.dCommitTrans

Set oGraba = Nothing

End Sub

Private Sub ImprimeCabeceraBoletaPago(ByVal psNomAge As String, ByVal psMoneda As String, ByVal psCuotasPagadas As String, _
        ByVal psFecha As String, ByVal psHora As String, ByVal psPersNombre As String, ByVal psCtaCod As String, Optional ByVal psCodCmac As String)

Dim oImpre As New COMFunciones.FCOMImpresion

        On Error GoTo ErrorImprimeCabeceraBoletaPago
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita

        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(71)
        vsCadImpresion = vsCadImpresion & Chr$(10)
        vsCadImpresion = vsCadImpresion & Chr$(10)
        vsCadImpresion = vsCadImpresion & Chr$(10)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CMI - CREDITOS" & Space(5) & oImpre.ImpreFormat("FIDEICOMISO", 51, 0), 53) & Trim(oImpre.ImpreFormat("CMI - CAPTACIONES" & Space(5) & oImpre.ImpreFormat("FIDEICOMISO", 51, 0), 53)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda, 53, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda, 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & Trim(oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & Chr$(10)
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description

End Sub

Public Sub ImprimeBoleta(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal pnCapitalPag As Double, ByVal pnInteresPag As Double, ByVal pnPortesPag As Double, _
            ByVal pnIntCompPag As Double, ByVal pnIntMorPag As Double, ByVal pnComVcdoPag As Double, _
            ByVal pnProtestoPag As Double, ByVal pnGastosPag As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psCodCmac As String)

Dim nFicSal As Integer
Dim oImpre As New COMFunciones.FCOMImpresion

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
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, psCuotasPagadas, psFecha, psHora, psPersNombre, psCtaCod, psCodCmac)

    'Cuerpo de la Boleta
If psCodCmac = "102" Then
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(71)
    vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapitalPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes :" & Space(13) & oImpre.ImpreFormat(pnInteresPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Portes:" & Space(15) & oImpre.ImpreFormat(pnPortesPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Comision Vcdo:" & Space(8) & oImpre.ImpreFormat(pnComVcdoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Protesto:" & Space(13) & oImpre.ImpreFormat(pnProtestoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastosPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & String(40, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True) & Chr$(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)
Else
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(71)
    vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapitalPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapitalPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes :" & Space(13) & oImpre.ImpreFormat(pnInteresPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes :" & oImpre.ImpreFormat(pnInteresPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Portes:" & Space(15) & oImpre.ImpreFormat(pnPortesPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Portes:" & Space(15) & oImpre.ImpreFormat(pnPortesPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Comision Vcdo:" & Space(8) & oImpre.ImpreFormat(pnComVcdoPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Comision Vcdo:" & Space(8) & oImpre.ImpreFormat(pnComVcdoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Protesto:" & Space(13) & oImpre.ImpreFormat(pnProtestoPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Protesto:" & Space(13) & oImpre.ImpreFormat(pnProtestoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastosPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastosPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True) & Chr$(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)
End If

    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal

    Exit Sub
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description

End Sub

Public Function ExtornoPagoFideicomiso(ByVal pnMovNroAnt As Long, ByVal psMovNroNew As String, ByVal psCodCta As String, ByVal pnMonto As Currency) As Boolean
Dim lsMovNroAnt As String
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
    
Dim oGeneral As New COMDConstSistema.DCOMGeneral
    ExtornoPagoFideicomiso = True
    
    Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(oGeneral.GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(oGeneral.GetFechaMov(psMovNroNew, True))

    If ldFechaAnt = ldFechaExt Then
        
        On Error GoTo dErrorExtornoPagoFideicomiso
        
        oGraba.dBeginTrans
        
        oGraba.dInsertMov psMovNroNew, gServExtCobFideicomiso, "Extorno Pago de Fideicomiso", gMovEstContabMovContable, gMovFlagExtornado
        lnMovNronew = oGraba.GetnMovNro(psMovNroNew)
        oGraba.dInsertMovRef lnMovNronew, pnMovNroAnt
        oGraba.dUpdateMov pnMovNroAnt
        oGraba.dUpdateFideicomisoMov pnMovNroAnt, gMovFlagExtornado
        lnCuotasPag = oGraba.dGetCuotasPag(pnMovNroAnt)
        oGraba.dUpdateFideicomiso psCodCta, -lnCuotasPag, -pnMonto 'Disminuye el total de monto pagado
        
        oGraba.dCommitTrans
                
    Else
        ExtornoPagoFideicomiso = False
    End If
        
    Exit Function
    
dErrorExtornoPagoFideicomiso:
    ExtornoPagoFideicomiso = False
    Err.Raise Err.Number, "Error en <<Error en ExtornoPagoFideicomiso >> ", Err.Description

End Function
'FONCODES
'CRSF - 24/06
Public Function ExtornoPagoFoncodes(ByVal pnMovNroAnt As Long, ByVal psMovNroNew As String, ByVal psCodCta As String, ByVal pnMonto As Currency) As Boolean
Dim lsMovNroAnt As String
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
Dim oGeneral As New COMDConstSistema.DCOMGeneral

    ExtornoPagoFoncodes = True
    
    Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(oGeneral.GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(oGeneral.GetFechaMov(psMovNroNew, True))

    If ldFechaAnt = ldFechaExt Then
        
        On Error GoTo dErrorExtornoPagoFoncodes
        
        oGraba.dBeginTrans
        
        oGraba.dInsertMov psMovNroNew, gServExtCobFoncodes, "Extorno Pago de Foncodes", gMovEstContabMovContable, gMovFlagExtornado
        lnMovNronew = oGraba.GetnMovNro(psMovNroNew)
        oGraba.dInsertMovRef lnMovNronew, pnMovNroAnt
        oGraba.dUpdateMov pnMovNroAnt
        oGraba.dUpdateFoncodesMov pnMovNroAnt, gMovFlagExtornado
        lnCuotasPag = oGraba.dGetCuotasPag(pnMovNroAnt)
        oGraba.dUpdateFONCODES psCodCta, -lnCuotasPag, -pnMonto 'Disminuye el total de monto pagado
        
        oGraba.dCommitTrans
                
    Else
        ExtornoPagoFoncodes = False
    End If
        
    Exit Function
    
dErrorExtornoPagoFoncodes:
    ExtornoPagoFoncodes = False
    Err.Raise Err.Number, "Error en <<Error en ExtornoPagoFoncodes>> ", Err.Description

End Function
'PLAN BICI
' CRSF - 24/06
Public Function ExtornoPagoPlanBici(ByVal pnMovNroAnt As Long, ByVal psMovNroNew As String, ByVal psCodCta As String, ByVal pnMonto As Currency) As Boolean
Dim lsMovNroAnt As String
Dim oGraba As COMDCaptaServicios.DCOMCaptaFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
Dim oGeneral As New COMDConstSistema.DCOMGeneral

    ExtornoPagoPlanBici = True
    
    Set oGraba = New COMDCaptaServicios.DCOMCaptaFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(oGeneral.GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(oGeneral.GetFechaMov(psMovNroNew, True))

    If ldFechaAnt = ldFechaExt Then
        
        On Error GoTo dErrorExtornoPagoPlanBici
        
        oGraba.dBeginTrans
        
        oGraba.dInsertMov psMovNroNew, gServExtCobPlanBici, "Extorno Pago de Plan Bici", gMovEstContabMovContable, gMovFlagExtornado
        lnMovNronew = oGraba.GetnMovNro(psMovNroNew)
        oGraba.dInsertMovRef lnMovNronew, pnMovNroAnt
        oGraba.dUpdateMov pnMovNroAnt
        oGraba.dUpdatePlanBiciMov pnMovNroAnt, gMovFlagExtornado
        lnCuotasPag = oGraba.dGetCuotasPagPlan(pnMovNroAnt)
        oGraba.dUpdatePLANBICI psCodCta, -lnCuotasPag, -pnMonto 'Disminuye el total de monto pagado
        
        oGraba.dCommitTrans
                
    Else
        ExtornoPagoPlanBici = False
    End If
        
    Exit Function
    
dErrorExtornoPagoPlanBici:
    ExtornoPagoPlanBici = False
    Err.Raise Err.Number, "Error en <<Error en ExtornoPagoplanBici>> ", Err.Description

End Function
