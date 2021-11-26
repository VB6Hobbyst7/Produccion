VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapFideicomiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsCadImpresion As String

Public Sub CapPagoFideicomiso(ByVal psCodCta As String, ByVal pmPlanPagos As Variant, ByVal pnMontoPago As Currency, _
                              ByVal pnCuotasPag As Integer, ByVal pnTipoPago As Integer, ByVal psMovNro As String, _
                              ByVal pnTipoCalend As Integer, Optional ByVal pnCuotaPagPreJud As Integer)
            
Dim oGraba As DCapFideicomiso
Dim lnMovNro As String
Dim I As Integer

Set oGraba = New DCapFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobFideicomiso, "Pago de Fideicomiso", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For I = 1 To pnCuotasPag
       
        If pmPlanPagos(I, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(I, 1), pmPlanPagos(I, 2)
        End If
        
        If pmPlanPagos(I, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(I, 1), pmPlanPagos(I, 3)
        End If
        
        If pmPlanPagos(I, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(I, 1), pmPlanPagos(I, 4)
        End If
        
        If pmPlanPagos(I, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(I, 1), pmPlanPagos(I, 5)
        End If
        
        If pmPlanPagos(I, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(I, 1), pmPlanPagos(I, 6)
        End If
        
        If pmPlanPagos(I, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(I, 1), pmPlanPagos(I, 7)
        End If
        
        If pmPlanPagos(I, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(I, 1), pmPlanPagos(I, 8)
        End If
        
        If pmPlanPagos(I, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(I, 1), pmPlanPagos(I, 9)
        End If
        
        oGraba.dInsertFideicomisoMov psCodCta, pnTipoPago, pmPlanPagos(I, 1), pmPlanPagos(I, 2), pmPlanPagos(I, 3), pmPlanPagos(I, 6), _
                                     pmPlanPagos(I, 4), pmPlanPagos(I, 5), pmPlanPagos(I, 7), pmPlanPagos(I, 8), pmPlanPagos(I, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next I
    
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
            
Dim oGraba As DCapFideicomiso
Dim lnMovNro As String
Dim I As Integer

Set oGraba = New DCapFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobFoncodes, "Pago de FONCODES", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For I = 1 To pnCuotasPag
       
        If pmPlanPagos(I, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(I, 1), pmPlanPagos(I, 2)
        End If
        
        If pmPlanPagos(I, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(I, 1), pmPlanPagos(I, 3)
        End If
        
        If pmPlanPagos(I, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(I, 1), pmPlanPagos(I, 4)
        End If
        
        If pmPlanPagos(I, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(I, 1), pmPlanPagos(I, 5)
        End If
        
        If pmPlanPagos(I, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(I, 1), pmPlanPagos(I, 6)
        End If
        
        If pmPlanPagos(I, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(I, 1), pmPlanPagos(I, 7)
        End If
        
        If pmPlanPagos(I, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(I, 1), pmPlanPagos(I, 8)
        End If
        
        If pmPlanPagos(I, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(I, 1), pmPlanPagos(I, 9)
        End If
        
        oGraba.dInsertFONCODESMov psCodCta, pnTipoPago, pmPlanPagos(I, 1), pmPlanPagos(I, 2), pmPlanPagos(I, 3), pmPlanPagos(I, 6), _
                                     pmPlanPagos(I, 4), pmPlanPagos(I, 5), pmPlanPagos(I, 7), pmPlanPagos(I, 8), pmPlanPagos(I, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next I
    
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
            
Dim oGraba As DCapFideicomiso
Dim lnMovNro As String
Dim I As Integer

Set oGraba = New DCapFideicomiso
        
    oGraba.dBeginTrans
    
    oGraba.dInsertMov psMovNro, gServCobPlanBici, "Pago de PLAN BICI", gMovEstContabMovContable, gMovFlagVigente
    
    lnMovNro = oGraba.GetnMovNro(psMovNro)
    
    'Inserta en MovServicios
    oGraba.dInsertMovServicios lnMovNro, psCodCta, "", pnMontoPago, Mid(psCodCta, 4, 2), gMovFlagVigente
        
    For I = 1 To pnCuotasPag
       
        If pmPlanPagos(I, 2) > 0 Then   'Capital
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidCapital, pmPlanPagos(I, 1), pmPlanPagos(I, 2)
        End If
        
        If pmPlanPagos(I, 3) > 0 Then   'Interes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidInteres, pmPlanPagos(I, 1), pmPlanPagos(I, 3)
        End If
        
        If pmPlanPagos(I, 4) > 0 Then   'IntComp
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidIntComp, pmPlanPagos(I, 1), pmPlanPagos(I, 4)
        End If
        
        If pmPlanPagos(I, 5) > 0 Then   'Mora
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidMora, pmPlanPagos(I, 1), pmPlanPagos(I, 5)
        End If
        
        If pmPlanPagos(I, 6) > 0 Then   'Portes
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidPortes, pmPlanPagos(I, 1), pmPlanPagos(I, 6)
        End If
        
        If pmPlanPagos(I, 7) > 0 Then   'Protesto
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidProtesto, pmPlanPagos(I, 1), pmPlanPagos(I, 7)
        End If
        
        If pmPlanPagos(I, 8) > 0 Then   'Comision Vencido
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidComVcdo, pmPlanPagos(I, 1), pmPlanPagos(I, 8)
        End If
        
        If pmPlanPagos(I, 9) > 0 Then   'Gastos
            oGraba.dInsertMovServiciosDet psCodCta, lnMovNro, gConcFidGastos, pmPlanPagos(I, 1), pmPlanPagos(I, 9)
        End If
        
        oGraba.dInsertPLANBICIMov psCodCta, pnTipoPago, pmPlanPagos(I, 1), pmPlanPagos(I, 2), pmPlanPagos(I, 3), pmPlanPagos(I, 6), _
                                     pmPlanPagos(I, 4), pmPlanPagos(I, 5), pmPlanPagos(I, 7), pmPlanPagos(I, 8), pmPlanPagos(I, 9), _
                                     lnMovNro, gMovFlagVigente
        
    Next I
    
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
        vsCadImpresion = vsCadImpresion & ImpreFormat("CMI - CREDITOS" & Space(5) & ImpreFormat("FIDEICOMISO", 51, 0), 53) & Trim(ImpreFormat("CMI - CAPTACIONES" & Space(5) & ImpreFormat("FIDEICOMISO", 51, 0), 53)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda, 53, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda, 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)
        vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & Chr$(10)
        vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & Chr$(10)
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
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapitalPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes :" & Space(13) & ImpreFormat(pnInteresPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Portes:" & Space(15) & ImpreFormat(pnPortesPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Comision Vcdo:" & Space(8) & ImpreFormat(pnComVcdoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Protesto:" & Space(13) & ImpreFormat(pnProtestoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastosPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & String(40, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True) & Chr$(10)
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)
Else
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(71)
    vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapitalPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Capital :" & Space(13) & ImpreFormat(pnCapitalPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes :" & Space(13) & ImpreFormat(pnInteresPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes :" & ImpreFormat(pnInteresPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Portes:" & Space(15) & ImpreFormat(pnPortesPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Portes:" & Space(15) & ImpreFormat(pnPortesPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes Compensatorio:" & ImpreFormat(pnIntCompPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Mora :" & Space(16) & ImpreFormat(pnIntMorPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Comision Vcdo:" & Space(8) & ImpreFormat(pnComVcdoPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Comision Vcdo:" & Space(8) & ImpreFormat(pnComVcdoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Protesto:" & Space(13) & ImpreFormat(pnProtestoPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Protesto:" & Space(13) & ImpreFormat(pnProtestoPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastosPag, 12, 2)
    vsCadImpresion = vsCadImpresion & Space(16) & "Gastos :" & Space(14) & ImpreFormat(pnGastosPag, 12, 2) & Chr$(10)
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & Chr$(10)
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapitalPag + pnInteresPag + pnPortesPag + pnIntCompPag + pnIntMorPag + pnComVcdoPag + pnProtestoPag + pnGastosPag, 12, 2, True) & Chr$(10)
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
Dim oGraba As DCapFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
    
    ExtornoPagoFideicomiso = True
    
    Set oGraba = New DCapFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(GetFechaMov(psMovNroNew, True))

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
Dim oGraba As DCapFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
    
    ExtornoPagoFoncodes = True
    
    Set oGraba = New DCapFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(GetFechaMov(psMovNroNew, True))

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
Dim oGraba As DCapFideicomiso
Dim ldFechaAnt As Date
Dim ldFechaExt As Date
Dim lnMovNronew As Long
Dim lnCuotasPag As Integer
    
    ExtornoPagoPlanBici = True
    
    Set oGraba = New DCapFideicomiso
    
    lsMovNroAnt = oGraba.GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(GetFechaMov(psMovNroNew, True))

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



