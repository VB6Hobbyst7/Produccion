VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCaptaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Genera los registros de Firmas, las cartas de contrato y las constancias de
'plazo fijo, saca las plantillas de la base de datos de acuerdo a un codigo
'que se le envía
Option Base 0
Option Explicit

Private lsNomAge As String
Private lsEmpresa As String
Private ldFecSis As Date
Private ldFecRepo As Date
Private lnRango As Integer
Private lsAgeCod As String
Private gsInstCmac As String
Private laMovDia() As String

Private lsRTFDat As String
Private i As Integer
Private lnTotA02 As Currency
Private lnTotA03 As Currency
Private lnTotA04 As Currency
Private lnTotA05 As Currency
Private lnTotA06 As Currency
Private lnTotA07 As Currency
Private lnTotA08 As Currency
Private lnTotA09 As Currency
Private lnTotA10 As Currency
Private lnTotA11 As Currency
Private lnTotA12 As Currency
Private lnTotA13 As Currency
Private lnTotA14 As Currency

Private lsRTFRfm As String
Private lsReporte As String
Private lsMoneda As String
Private lsRTFSdo As String
Private lsPerson As String
Dim VSQL As String
Dim SalA As String
Dim NumA As String

Dim SalC As String
Dim NumC As String

Dim NumP As String
Dim SalP As String

Dim CA As Currency
Dim CAE As Currency
Dim TB As Currency
Dim TBD As Currency

Dim ImpSoles As String
Dim k As Integer
Dim DHora(50) As String
Dim TC(100) As String
Dim TCE(100) As String
Dim TV(100) As String
Dim TVE(100) As String

Dim lnTipoPigno As Integer
Dim RegCmacStemp As Integer

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3A835D5E0029
    Private mlClassDebugID As Long
#End If

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        'mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3A835D5E028C
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3A835D5E00BF
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'dd/mm/yyyy
Public Function Reporte(psRep As String, psFecini As String, psFecFin As String, pnMontoIni As Currency, pnMontoFin As Currency, _
                        psNomAge As String, psEmpresa As String, pdFecSis As Date, psCodAge As String, _
                        Optional psUser As String = "XXX", Optional psTextoCarta As String = "", Optional pnTipoCambio As Double = 1, Optional lsEstadosCheques_ As String = "", Optional lsOptionsCheques_ As String = "", Optional lsOrden_ As String, Optional lsCheck_ As String, Optional lsLlamada_ As String = "0", Optional lsrecepcion_ As String = "0", Optional PSpersoneria As Integer) As String
    Dim lsCadena As String
    Dim lsRep As String
    Dim lsRepSub As String
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    lsNomAge = psNomAge
    lsEmpresa = psEmpresa
    ldFecSis = pdFecSis
    ldFecRepo = CDate(pdFecSis)
    lsAgeCod = psCodAge
    gsCodAge = psCodAge
    lnRango = 54
    gdFecSis = pdFecSis
    Dim gsCodProAC As String
    Dim gsCodProPF As String
    Dim gsCodProCTS As String
    
    
    If Len(Trim(psUser)) = 0 Then
        psUser = "XXX"
    End If
    
    gsCodAge = psCodAge
    gsCodProAC = Producto.gCapAhorros
    gsCodProPF = Producto.gCapPlazoFijo
    gsCodProCTS = Producto.gCapCTS
    
    lnTipoPigno = oConst.LeeConstSistema(101)

    lsRep = psRep
    lsRepSub = lsRep
    If lsRepSub = gCapRepDiaCapEstadAho Then 'ok
        lsCadena = lsCadena & LlenaEDAC(1, psFecini, psFecFin)
        lsCadena = lsCadena & LlenaEDAC(2, psFecini, psFecFin)
    ElseIf lsRepSub = gCapRepDiaCapEstadPF Then 'ok
        lsCadena = lsCadena & LlenaEDPF(1, psFecini, psFecFin)
        lsCadena = lsCadena & LlenaEDPF(2, psFecini, psFecFin)
    ElseIf lsRepSub = gCapRepDiaCapEstadCTS Then 'ok
        lsCadena = lsCadena & LlenaEDCTS(1, psFecini, psFecFin)
        lsCadena = lsCadena & LlenaEDCTS(2, psFecini, psFecFin)
    ElseIf lsRepSub = gCapRepDiaCapCtasMov Then 'ok
        lsCadena = lsCadena & ReporteTrasTotSM("MOVIMIENTOS DE AHORRO", False, psUser, Format(CDate(psFecini), "yyyymmdd"))
    ElseIf lsRepSub = gCapRepDiaCapSaldTpoCta Then 'ok
        lsCadena = lsCadena & SdosTipCta(lsAgeCod)
    ElseIf lsRepSub = gCapRepDiaCapEstratCta Then 'ok
        lsCadena = lsCadena & EstratDep(lsAgeCod)
    ElseIf lsRepSub = gCapRepDiaCapPFVenc Then 'ok
        lsCadena = lsCadena & RepPlazoFijovencimiento(CDate(psFecini), CDate(psFecFin), psNomAge, psEmpresa, pdFecSis, gMonedaNacional)
        lsCadena = lsCadena & Chr(12) & RepPlazoFijovencimiento(CDate(psFecini), CDate(psFecFin), psNomAge, psEmpresa, pdFecSis, gMonedaExtranjera)
    ElseIf lsRepSub = gCapRepDiaCapInact Then 'ok
        lsCadena = lsCadena & RepCtaIna
    ElseIf lsRepSub = gCapRepDiaCapConsInact Then 'ok
        lsCadena = lsCadena & ConsolidadoInactivas(CDate(psFecini), CDate(psFecFin))
    ElseIf lsRepSub = gCapRepDiaCapApert Then 'Cuentas Aperturadas en el día 'ok
        lsCadena = lsCadena & RepCtaApe(CDate(psFecini), CDate(psFecFin))
    ElseIf lsRepSub = gCapRepDiaCapCanc Then 'ok
        lsCadena = lsCadena & RepCtaCan(CDate(psFecini), CDate(psFecFin))
    ElseIf lsRepSub = gCapRepDiaCartaRenPF Then 'ok
        lsCadena = lsCadena & AvisoRenovacionPF(CDate(psFecini), psTextoCarta, psCodAge)
        
    ElseIf lsRepSub = gCapRepDiaServGirosApert Then  'GIRO APE 'ok
        lsCadena = lsCadena & ImprimeRepGiros(CDate(psFecini), CDate(psFecFin), "MSM", "GIROS REGISTRADOS")
    ElseIf lsRepSub = gCapRepDiaServGirosCanc Then 'GIRO CAN 'ok
        lsCadena = lsCadena & ImprimeRepGiros(CDate(psFecini), CDate(psFecFin), "MSM", "GIROS CANCELADOS", True)
    ElseIf lsRepSub = gCapRepDiaServConvCob Then  'SERVICIOS
        lsCadena = lsCadena & RepCtaCan(CDate(psFecini), CDate(psFecFin))
        
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoPN Then 'PERSONAS NATURALES CMCPL 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 1)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 1)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoPJSFL Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoPJCFL Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoCMAC Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoCRAC Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoFoncodes Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLFONCODES, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLFONCODES, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoCooperativa Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaAhoEdipyme Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaExtranjera, pdFecSis)
   ElseIf lsRepSub = "280632" Then 'ok
        lsCadena = lsCadena & GeneraRepInac(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, Format(gdFecSis, "mm/dd/yyyy"), 1)
        lsCadena = lsCadena & GeneraRepInac(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, Format(gdFecSis, "mm/dd/yyyy"), 1)
   ElseIf lsRepSub = "280633" Then 'ok
        lsCadena = lsCadena & GeneraRepInac(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, Format(gdFecSis, "mm/dd/yyyy"), 2)
        lsCadena = lsCadena & GeneraRepInac(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, Format(gdFecSis, "mm/dd/yyyy"), 2)
   ElseIf lsRepSub = "280634" Then 'ok
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, 4)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, 4)
   ElseIf lsRepSub = "280635" Then 'ok
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, 5)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, 5)
   ElseIf lsRepSub = "280636" Then 'ok
        lsAgeCod = "X4"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis)
   ElseIf lsRepSub = "280637" Then 'ok
        lsAgeCod = "X2"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis)
   ElseIf lsRepSub = "280639" Then
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaNacional, pdFecSis, 5)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis, 5)
   
    
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFPN Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFPJSFL Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFPJCFL Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFCMAC Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFCRAC Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFFoncodes Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLFONCODES, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLFONCODES, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFCooperativa Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFEdpyme Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFPNPJSFL Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaNacional, pdFecSis, 5)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis, 5)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFNascaPalpa Then 'ok
        lsAgeCod = "X4"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaPFCaneteMala Then 'ok
        lsAgeCod = "X2"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaNacional, pdFecSis)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaExtranjera, pdFecSis)
   ElseIf lsRepSub = gCapRepMensBloqCapPF Then 'ok
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaNacional, 4)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaExtranjera, 4)
   ElseIf lsRepSub = "280631" Then 'ok
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaNacional, 5)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapPlazoFijo, PSpersoneria, Moneda.gMonedaExtranjera, 5)
    
    
    ElseIf lsRepSub = gCapRepMensCapSaldCtaCTS Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 1)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 1)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaCTSConvenio Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 2)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 2)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaCTSExternos Then 'ok
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 3)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 3)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaCTSNascaPalpa Then 'ok
        lsAgeCod = "X4"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 1)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 1)
    ElseIf lsRepSub = gCapRepMensCapSaldCtaCTSCaneteMala Then 'ok
        lsAgeCod = "X2"
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, pdFecSis, 1)
        lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, pdFecSis, 1)
   ElseIf lsRepSub = gCapRepMensBloqCapCTS Then
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, 1)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, 1)
   ElseIf lsRepSub = gCapRepMensBloqCapCTSConvenio Then
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, 2)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, 2)
   ElseIf lsRepSub = gCapRepMensBloqCapCTSExternos Then
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, 3)
        lsCadena = lsCadena & GeneraRepBloq(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, 3)
        
   ElseIf lsRepSub = gCapRepMensCapListGralCtas Then 'ok
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaExtranjera, psFecini, 0)
        lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaExtranjera, psFecini, 0)
       
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCMAC, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCRAC, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLCooperativa, Moneda.gMonedaExtranjera, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaNacional, psFecini, 1)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, PersPersoneria.gPersonaJurCFLEdpyme, Moneda.gMonedaExtranjera, psFecini, 1)
    
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurSFL, Moneda.gMonedaExtranjera, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaNacional, psFecini, 0)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, PersPersoneria.gPersonaJurCFL, Moneda.gMonedaExtranjera, psFecini, 0)
    
       lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, psFecini, 0, gsInstCmac)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, psFecini, 0, "")
       lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaNacional, psFecini, 0, gsInstCmac)
       lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, PersPersoneria.gPersonaNat, Moneda.gMonedaExtranjera, psFecini, 0, "")
    ElseIf lsRepSub = "V03" Then
        lsCadena = lsCadena & CuentasInactivas(Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin))
        lsCadena = lsCadena & CuentasInactivas(Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin))
    
    ElseIf lsRepSub = "280701" Then 'ok
        lsCadena = lsCadena & PlazosFijoCancelayMontos(Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
    ElseIf lsRepSub = "280702" Then 'ok
        lsCadena = lsCadena & PlazosFijoCancelayMontos(Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
    ElseIf lsRepSub = "280703" Then 'ok
        lsCadena = lsCadena & PlazosFijoCancelayMontosEspecial(Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
    ElseIf lsRepSub = "280704" Then 'ok
        lsCadena = lsCadena & PlazosFijoCancelayMontosEspecial(Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
    
    ElseIf lsRepSub = "280705" Then 'ok
        lsCadena = lsCadena & ReportesExtornos(psCodAge, "REPORTES DE CUENTAS EXTORNADAS", Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin, psNomAge)
    
    ElseIf lsRepSub = "280706" Then 'ok
        'ahorros
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE AHORRO APERTURADAS SIN OREDN DE PAGO ", Producto.gCapAhorros, 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE AHORRO APERTURADAS CON OREDN DE PAGO ", Producto.gCapAhorros, 1, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE AHORRO APERTURADAS SIN OREDN DE PAGO ", Producto.gCapAhorros, 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE AHORRO APERTURADAS CON OREDN DE PAGO ", Producto.gCapAhorros, 1, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        
        'plazo fijo
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE PLAZO FIJO APERTURADAS ", Producto.gCapPlazoFijo, 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE PLAZO FIJO APERTURADAS ", Producto.gCapPlazoFijo, 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        
        'Cts
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE CTS APERTURADAS ", Producto.gCapCTS, 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasAperturadas(psCodAge, "CUENTAS DE CTS APERTURADAS ", Producto.gCapCTS, 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        
    ElseIf lsRepSub = "280707" Then 'ok
        'Canc. ahorros
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE AHORRO CANCELADAS SIN OREDN DE PAGO ", "232", 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE AHORRO CANCELADAS CON OREDN DE PAGO ", "232", 1, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE AHORRO CANCELADAS SIN OREDN DE PAGO ", "232", 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE AHORRO CANCELADAS CON OREDN DE PAGO ", "232", 1, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        
        'Canc. plazo fijo
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE PLAZO FIJO CANCELADAS ", "233", 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE PLAZO FIJO CANCELADAS ", "233", 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        
        'Canc. Cts
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE CTS CANCELADAS ", "233", 0, Moneda.gMonedaNacional, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
        lsCadena = lsCadena & CuentasCanceladas(psCodAge, "CUENTAS DE CTS CANCELADAS ", "233", 0, Moneda.gMonedaExtranjera, CDate(psFecini), CDate(psFecFin), pnMontoIni, pnMontoFin)
    
    ElseIf lsRepSub = "280213" Then 'ok
        lsCadena = lsCadena & ProtocoloOperaciones("Protocolo del Dia Soles", 0, 1, psNomAge, psEmpresa, pdFecSis, gMonedaNacional, psUser, , CDate(psFecini), psCodAge, lsOrden_, lsCheck_)
    ElseIf lsRepSub = "280214" Then 'ok
        lsCadena = lsCadena & ProtocoloOperaciones("Protocolo del Dia Dolares", 0, 1, psNomAge, psEmpresa, pdFecSis, gMonedaExtranjera, psUser, , CDate(psFecini), psCodAge, lsOrden_, lsCheck_)
    ElseIf lsRepSub = "280215" Then 'ok
        lsCadena = lsCadena & ComVtaME(CDate(psFecini), CDate(psFecFin), psCodAge)
    ElseIf lsRepSub = "280216" Then 'ok
        lsCadena = lsCadena & InfCajaGeneral()
        
    ElseIf lsRepSub = "280217" Then 'ok
        Dim bHoy As Boolean
        bHoy = False
        If pdFecSis = CDate(psFecini) Then bHoy = True
        lsCadena = lsCadena & SaldosDiarios(CDate(psFecini), psCodAge, psNomAge, bHoy, psEmpresa)
    
    ElseIf lsRepSub = "280708" Then 'ok
        lsCadena = lsCadena & GetRepCapNMejoresClientes(CLng(pnMontoIni), pnTipoCambio, psCodAge, psEmpresa, psNomAge, pdFecSis)
    ElseIf lsRepSub = "280224" Then
        lsCadena = lsCadena & RepCtaInmobilizada(CDate(psFecini), CDate(psFecFin))
    ElseIf lsRepSub = "280225" Then
        lsCadena = lsCadena & ImpLavDineroD(CDate(psFecini), CDate(psFecFin), psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280227" Then
        lsCadena = lsCadena & ImpLavDineroAD(CDate(psFecini), psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280228" Then
        lsCadena = lsCadena & ImpLavDineroM(CDate(psFecini), CDate(psFecFin), psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280219" Then 'ok
        lsCadena = lsCadena & SobranteFaltante(CDate(psFecini), CDate(psFecFin), psCodAge)   ' , psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280220" Then 'ok
        lsCadena = lsCadena & GetOtrasOperaciones(CDate(psFecini), CDate(psFecFin), psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280230" Then 'Cheques x estado
    
        lsCadena = lsCadena & RepoChequesCompleto(psCodAge, CDate(psFecini), CDate(psFecFin), lsEstadosCheques_, lsOptionsCheques_, 1, psNomAge, Format(pdFecSis, "dd/MM/YYYY"))       ' , psEmpresa, psNomAge, psCodAge, pdFecSis)
        
        If Len(Trim(lsCadena)) > 0 Then
            lsCadena = lsCadena & Chr(12)
        End If
        
        lsCadena = lsCadena & RepoChequesCompleto(psCodAge, CDate(psFecini), CDate(psFecFin), lsEstadosCheques_, lsOptionsCheques_, 2, psNomAge, Format(pdFecSis, "dd/MM/YYYY"))     ' , psEmpresa, psNomAge, psCodAge, pdFecSis)
        
    ElseIf lsRepSub = "280223" Then
        lsCadena = lsCadena & GetListaOPSolicitadas(CDate(psFecini), CDate(psFecFin), psEmpresa, psNomAge, psCodAge, pdFecSis)
    ElseIf lsRepSub = "280222" Then 'ok
        If lsLlamada_ = "1" Then
            lsCadena = lsCadena & GetListaLlamadasCMACS(CDate(psFecini), CDate(psFecini), psEmpresa, psNomAge, psCodAge, pdFecSis)
        End If
        
        If lsrecepcion_ = "1" Then
            If lsLlamada_ = "1" Then
                lsCadena = lsCadena & Chr(12)
            End If
            lsCadena = lsCadena & GetListaRecepcionCMACS(CDate(psFecini), CDate(psFecini), psEmpresa, psNomAge, psCodAge, pdFecSis)
        End If
    End If
    
    Reporte = lsCadena
End Function

Public Function ReporteTrasTotSM(psTitulo As String, Optional pbFinPg As Boolean = True, Optional psCodUser As String = "XXX", Optional psFecha As String = "") As String
        '*********CODIGO ORIGINAL
        
    'On Error GoTo ERROR
    Dim lsCadena As String
    Dim lsAdd As String
    lsCadena = ""
    

    
    lsAdd = IIf(psFecha = "", "", " Del " & psFecha)
    ReporteMovAC lsCadena, "Detalle de Operaciones Ahorro - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPF lsCadena, "Detalle de Operaciones Plazo Fijo - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCTS lsCadena, "Detalle de Operaciones CTS - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCreditos lsCadena, "Detalle de Operaciones Creditos - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPignoraticio lsCadena, "Detalle de Operaciones Pignoraticio - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovJudicial lsCadena, "Detalle de Operaciones Judicial - Soles" & lsAdd, 0, 0, Moneda.gMonedaNacional, psCodUser, psFecha
    
    PonerSalto lsCadena
    ReporteMovAC lsCadena, "Detalle de Operaciones Ahorro - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPF lsCadena, "Detalle de Operaciones Plazo Fijo - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCTS lsCadena, "Detalle de Operaciones CTS - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCreditos lsCadena, "Detalle de Operaciones Creditos - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPignoraticio lsCadena, "Detalle de Operaciones Pignoraticio - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovJudicial lsCadena, "Detalle de Operaciones Judicial - Dolares" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    
    PonerSalto lsCadena
    ReporteCompraVenta lsCadena, "Detalle de Operaciones Compra Venta" & lsAdd, 0, 0, Moneda.gMonedaExtranjera, psCodUser, psFecha
    
    If pbFinPg Then lsCadena = lsCadena & Chr(12)

    'QuitarSalto lsCadena

    ReporteTrasTotSM = lsCadena
    
    Exit Function
Error:
    MsgBox "Error Realizando el Reporte de Movimientos de Cuentas: " & Err.Description, vbCritical, "Aviso"
    '************CODIGO ORIGINAL
End Function




'************************************************************************
'Ahorros
'********************************************************************************************
Public Sub ReporteMovAC(psCadenaRep As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, _
                        Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double
    Dim lbBan As Boolean
    
    Dim lnTotDep As Double
    Dim lnTotRet As Double
    
    Dim lsCabecera As String
    
    Dim i As Long
    Dim lsTEM As String
    Dim lsAdd As String
    lnTotRet = 0
    lnTotDep = 0
    
    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))
    
    If psCodUser = "XXX" Then
        lsCabecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, Str(psMoneda))
    Else
        lsCabecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, Str(psMoneda))
    End If
    
    lsTEM = psCadenaRep
    
    
    psCadenaRep = ""
    
    pnPagina = pnPagina - 1
    
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas de Ahorro en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoApeEfec, , True))           'OK
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas de Ahorro con Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoApeCheque, , True))          'OK
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas por Desembolsos" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoApeDesembolso, , True))          'OK
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas por Transferencia" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoApeTreansferencia, , True))    'OK
    Call ReporteMovACDet(psCadenaRep, "Movimientos Sin Dinero en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoMovSinEfec, , True))     'OK
    Call MovimientoEfectivoAC(psCadenaRep, "Movimientos en Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoMovEfec, , True))
    Call ReporteMovACDet(psCadenaRep, "Cancelación de Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoCancelacion, , True))
    Call ReporteMovACDet(psCadenaRep, "Deposito s con Cheque en Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovAhoDepCheque, , True))     'OK
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadenaRep = lsCabecera + psCadenaRep
        QuitarSalto psCadenaRep
        psCadenaRep = psCadenaRep + Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
        psCadenaRep = psCadenaRep + Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";57; ;46;", pnItem)
        psCadenaRep = psCadenaRep & Chr(12)
    End If

    pnItem = 0
    pnPagina = 0
    lnTotDep = 0
    lnTotRet = 0
    
'    Call MovimientoExternoAC(psCadenaRep, "Cuentas de -Ahorro- Movidas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoDepEfec & "','" & CaptacOperacion.gAhoRetEfec & "','" & CaptacOperacion.gAhoRetOP & "','" & CaptacOperacion.gAhoRetOPCanje & "','" & CaptacOperacion.gAhoRetNC & "','" & CaptacOperacion.gAhoDepNA & "','" & CaptacOperacion.gAhoOPEmision & "','" & CaptacOperacion.gAhoRetOPCert & "','" & CaptacOperacion.gAhoRetOPCertCanje & "','" & CaptacOperacion.gAhoOPCertificacion & "','" & CaptacOperacion.gAhoRetAnulChq)
'    Call MovimientoExternoAC(psCadenaRep, "Cuentas de -Ahorro- Canceladas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoExtCancAct & "','" & CaptacOperacion.gAhoCancNCAct & "','" & CaptacOperacion.gAhoCancNCInact)
'    Call MovimientoExternoAC(psCadenaRep, "Cuentas de -Ahorro- Movidas en Otras Agencias" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoDepEfec & "','" & CaptacOperacion.gAhoRetEfec & "','" & CaptacOperacion.gAhoRetOP & "','" & CaptacOperacion.gAhoRetOPCanje & "','" & CaptacOperacion.gAhoRetNC & "','" & CaptacOperacion.gAhoDepNA & "','" & CaptacOperacion.gAhoOPEmision & "','" & CaptacOperacion.gAhoRetOPCert & "','" & CaptacOperacion.gAhoRetOPCertCanje & "','" & CaptacOperacion.gAhoOPCertificacion & "','" & CaptacOperacion.gAhoRetAnulChq, True)
'    'Call MovimientoTramiteACCMAC(psCadenaRep, "Cuentas de -Ahorro- Movidas en CMACs" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha)
    
    psCadenaRep = lsTEM + psCadenaRep
    lsTEM = ""
End Sub

Public Function ReporteMovACDet(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnTDepT As Double
    Dim lnTRetT As Double
    Dim lnValor As Double
    Dim lbBan As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    lbBan = False
    
    lnTDep = 0
    lnTRet = 0
    lnTDepT = 0
    lnTRetT = 0
    
    oCon.AbreConexion
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & "  And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
     lbBan = True
     psCadena = psCadena & CabeceraPagina(psTitulo & "- Sin Orden de Pago -", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
     psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usuario;10;Hora;17;Sald.Disp.;22;TInt;11;", pnItem)
    
     While Not RegTrandiaria.EOF
         
         If IsNull(RegTrandiaria!cNumDoc) Then
             lsDoc = " "
         Else
             lsDoc = RegTrandiaria!cNumDoc
         End If
         
         If RegTrandiaria!nMonTran >= 0 Then
             lnTDep = lnTDep + RegTrandiaria!nMonTran
         Else
             lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
         End If
         
         If IsNull(RegTrandiaria!nSaldDispAC) Then
             lnValor = 0
         Else
             lnValor = RegTrandiaria!nSaldDispAC
         End If
         
         psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                             & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                             & Space(19 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                             & Space(10 - Len(Right(RegTrandiaria!cMovNro, 4))) & Right(RegTrandiaria!cMovNro, 4) _
                             & Space(17 - Len(Mid(RegTrandiaria!cMovNro, 9, 6)) - 2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                             & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                             & Space(11 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                              
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & Chr(12)
             psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
             psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usuario;10;Hora;17;Sald.Disp.;22;TInt;11;", pnItem)
         End If
         
         RegTrandiaria.MoveNext
       Wend
       
       psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
           
        lnTDepT = lnTDepT + lnTDep
        lnTRetT = lnTRetT + lnTRet
        
        lnTDep = 0
        lnTRet = 0

        RegTrandiaria.Close
    End If
    Set RegTrandiaria = Nothing

'Orden de Pago
      
    
    If psCodUser = "XXX" Then
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2)  " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
     If pnItem < lnRango And lbBan Then
        psCadena = psCadena & Chr(12)
     End If
    
     psCadena = psCadena & CabeceraPagina(psTitulo & "- Con de Orden Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
     psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usuario;10;Hora;17;Sald.Disp.;22;TInt;11;", pnItem)
    
     While Not RegTrandiaria.EOF
         
         If IsNull(RegTrandiaria!cNumDoc) Then
             lsDoc = " "
         Else
             lsDoc = RegTrandiaria!cNumDoc
         End If
         
         If RegTrandiaria!nMonTran >= 0 Then
             lnTDep = lnTDep + RegTrandiaria!nMonTran
         Else
             lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
         End If
         
         If IsNull(RegTrandiaria!nSaldDispAC) Then
             lnValor = 0
         Else
             lnValor = RegTrandiaria!nSaldDispAC
         End If
         
         psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                            & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                            & Space(19 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                            & Space(10 - Len(Right(RegTrandiaria!cMovNro, 4))) & Right(RegTrandiaria!cMovNro, 4) _
                            & Space(17 - Len(Left(RegTrandiaria!cMovNro, 6)) - 2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                            & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                            & Space(11 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                            
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & Chr(12)
             psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
             psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usuario;10;Hora;17;Sald.Disp.;22;TInt;11;", pnItem)
         End If
         
         RegTrandiaria.MoveNext
       Wend
       
       psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";43; ;59;", pnItem)
           
        lnTDepT = lnTDepT + lnTDep
        lnTRetT = lnTRetT + lnTRet
        
        lnTDep = 0
        lnTRet = 0

        RegTrandiaria.Close
    End If
    
    Set RegTrandiaria = Nothing
    
    If lnTRetT <> 0 Or lnTDepT <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15;" & Trim(Format(lnTDepT, "###,##0.00")) & ";43; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRetT
        pnTotDep = pnTotDep + lnTDepT
        
        psCadena = psCadena & Chr(12)
    End If
End Function

Public Function MovimientoEfectivoAC(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lsDep As String
    Dim lsRet As String
    
    Dim lnTDep As Double
    Dim lnTRet As Double
    
    Dim lnTDepT As Double
    Dim lnTRetT As Double
    Dim lnValor As Double
    Dim lnOpe As Double
    Dim lbBan As Boolean
    
    Dim i As Long
    Dim sqlAdd As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lnTRet = 0
    lnTDep = 0
    
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' "
    End If
    
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.BOF And RegTrandiaria.EOF) Then
        lbBan = True
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
        
        While Not RegTrandiaria.EOF
            
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = RegTrandiaria!cNumDoc
            End If
                        
            If Left(RegTrandiaria!cCodOpe, 4) = Left(CaptacOperacion.gAhoDepEfec, 4) Or RegTrandiaria!cCodOpe = CaptacOperacion.gAhoTransAbono Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnOpe = RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnOpe = RegTrandiaria!nMonTran * -1
            End If
            
            If IsNull(RegTrandiaria!nSaldCnt) Then
                lnValor = 0
            Else
                lnValor = RegTrandiaria!nSaldCnt
            End If
            
            psCadena = psCadena & Space(20 - Len(Trim(RegTrandiaria!cCodCta))) & Trim(RegTrandiaria!cCodCta) _
                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                               & Space(18 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
                               & Space(18 - Len(lsDep)) & lsDep _
                               & Space(20 - Len(lsRet)) & lsRet _
                               & Space(2) & Right(RegTrandiaria!cMovNro, 4) _
                               & Space(2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
        
        psCadena = psCadena & Encabezado("Resumen;15;" & Format(lnTDep, "###,##0.00") & ";57;" & Format(lnTRet, "###,##0.00") & ";20; ;26;", pnItem)
        
        lnTRetT = lnTRet
        lnTDepT = lnTDep
            
        RegTrandiaria.Close
    End If
    
    'Ordenes de Pago
    lnTRet = 0
    lnTDep = 0
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.BOF And RegTrandiaria.EOF) Then
        If pnItem < lnRango And lbBan Then
            psCadena = psCadena & Chr(12)
        End If
        
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
        
        While Not RegTrandiaria.EOF
            
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = RegTrandiaria!cNumDoc
            End If
            
            
            If Left(RegTrandiaria!cCodOpe, 4) = Left(CaptacOperacion.gAhoDepEfec, 4) Or RegTrandiaria!cCodOpe = CaptacOperacion.gAhoTransAbono Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnOpe = RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnOpe = RegTrandiaria!nMonTran * -1
            End If
            
            If IsNull(RegTrandiaria!nSaldCnt) Then
                lnValor = 0
            Else
                lnValor = RegTrandiaria!nSaldCnt
            End If
            
            psCadena = psCadena & Space(20 - Len(Trim(RegTrandiaria!cCodCta))) & Trim(RegTrandiaria!cCodCta) _
                                & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                                & Space(18 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
                                & Space(18 - Len(lsDep)) & lsDep _
                                & Space(20 - Len(lsRet)) & lsRet _
                                & Space(2) & Right(RegTrandiaria!cMovNro, 4) _
                                & Space(2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
        
        lnTRetT = lnTRetT + lnTRet
        lnTDepT = lnTDepT + lnTDep
        
        psCadena = psCadena & Encabezado("Resumen;15;" & Format(lnTDep, "###,##0.00") & ";57;" & Format(lnTRet, "###,##0.00") & ";20; ;26;", pnItem)
        
        RegTrandiaria.Close
        Set RegTrandiaria = Nothing
    End If
    
    If lnTRetT <> 0 Or lnTDepT <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15;" & Format(lnTDepT, "###,##0.00") & ";57;" & Format(lnTRetT, "###,##0.00") & ";20; ;26;", pnItem)
        pnTotRet = pnTotRet + lnTRetT
        pnTotDep = pnTotDep + lnTDepT
        
        psCadena = psCadena & Chr(12)
    End If
    
End Function

Private Sub PonerSalto(rtf As String)
    If rtf = "" Then Exit Sub
    If Right(rtf, 1) <> Chr(12) Then
        rtf = rtf & Chr(12)
    End If
End Sub

Private Sub QuitarSalto(rtf As String)
    If Right(rtf, 1) = Chr(12) Then
        rtf = Mid(rtf, 1, Len(rtf) - 1)
    End If
End Sub

'**********************************************************
'CMAC's
'**********************************************************
'Movimiento que se realiza con cuentas de esta agencia pero
'por otro servidor
'Public Function MovimientoExternoACCmact(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pdFI As Date, pdFF As Date, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psTipoOpe As String) As String
'    Dim SQL1 As String
'    Dim RegTrandiaria As New ADODB.Recordset
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lsDep As String
'    Dim lsRet As String
'    Dim lsCodUsu As String
'
'    Dim lnTDep As Double
'    Dim lnTRet As Double
'    Dim lnTDepTT As Double
'    Dim lnTRetTT As Double
'
'    Dim lbBan As Boolean
'
'    Dim I As Long
'
'    Dim lsGru As String
'
'    Dim lnTDepAge As Double
'    Dim lnTRetAge As Double
'
'    Dim lnTDepCMAC As Double
'    Dim lnTRetCMAC As Double
'
'    Dim lsNomOpe As String * 40
'    Dim lsProd As String
'    Dim lsOP As String
'
'    lnTRet = 0
'    lnTDep = 0
'    lnTRetTT = 0
'    lnTDepTT = 0
'
'    lbBan = False
'
'    Dim lsFI As String
'    Dim lsFF As String
'
'    lsFI = Format(pdFI, gsFormatoFecha)
'    lsFF = Format(DateAdd("d", 1, pdFF), gsFormatoFecha)
'
'    If psCodUser = "XXX" Then
'        SQL1 = "select dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag,cCodCta,nMonTran,cCodAge,cCodUsu " _
'             & "from otrasope where ccodope in " _
'             & "(select ccodope from " & gcCentralCom & "Operacion where cGruProd in " _
'             & "(select cGruProd from " & gcCentralCom & "OpeGru where ccodtab = 'AA03' and cMovCaj = 'N' and cIngEgr = 'I')) and SubString(cCodCta,6,1) = '" & psMoneda & "' and cFlag is null and (dFecTran between '" & lsFI & "' and '" & lsFF & "') " _
'             & "Union " _
'             & "select dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag,cCodCta,nMonTran = nMonTran * -1 ,cCodAge,cCodUsu " _
'             & "from otrasope where ccodope in " _
'             & "(select ccodope from " & gcCentralCom & "Operacion where cGruProd in " _
'             & "(select cGruProd from " & gcCentralCom & "OpeGru where ccodtab = 'AA03' and cMovCaj = 'N' and cIngEgr = 'E')) and SubString(cCodCta,6,1) = '" & psMoneda & "' and cFlag is null and (dFecTran between '" & lsFI & "' and '" & lsFF & "') order by cCodAge, dFecTran"
'    Else
'        SQL1 = "select dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag,cCodCta,nMonTran,cCodAge,cCodUsu " _
'             & "from otrasope where ccodope in " _
'             & "(select ccodope from " & gcCentralCom & "Operacion where cGruProd in " _
'             & "(select cGruProd from " & gcCentralCom & "OpeGru where ccodtab = 'AA03' and cMovCaj = 'N' and cIngEgr = 'I')) and SubString(cCodCta,6,1) = '" & psMoneda & "' and cFlag is null and (dFecTran between '" & lsFI & "' and '" & lsFF & "') and cCodUsu = '" & psCodUser & "'" _
'             & "Union " _
'             & "select dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag,cCodCta,nMonTran = nMonTran * -1 ,cCodAge,cCodUsu " _
'             & "from otrasope where ccodope in " _
'             & "(select ccodope from " & gcCentralCom & "Operacion where cGruProd in " _
'             & "(select cGruProd from " & gcCentralCom & "OpeGru where ccodtab = 'AA03' and cMovCaj = 'N' and cIngEgr = 'E')) and SubString(cCodCta,6,1) = '" & psMoneda & "' and cFlag is null and (dFecTran between '" & lsFI & "' and '" & lsFF & "') and cCodUsu = '" & psCodUser & "' order by cCodAge, dFecTran"
'    End If
'
'    'RegTrandiaria.Open SQL1, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RSVacio(RegTrandiaria) Then
'        lbBan = False
'        'psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, psMoneda)
'        'psCadena = psCadena & Encabezado("NroDocCMAC;10;Tipo Operacion;40;Nro Cuenta;14;Deposito;17;Retiro;17;Usu;6;FechaHora;24;Pro.;4;OP;4;", pnItem)
'
'        lsGru = ""
'
'        lnTRetCMAC = 0
'        lnTDepCMAC = 0
'
'
'        While Not RegTrandiaria.EOF
'            lbBan = True
'            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
'                psCadena = psCadena & Encabezado("Total Agencia;15; ;18;" & Format(lnTDepAge, "###,##0.00") & ";50;" & Format(lnTRetAge, "###,##0.00") & ";17; ;36;", pnItem)
'
'                If Mid(lsGru, 1, 3) <> Mid(RegTrandiaria!cCodAge, 1, 3) Then
'                    psCadena = psCadena & Encabezado("Total CMAC;15; ;18;" & Format(lnTDepCMAC, "###,##0.00") & ";50;" & Format(lnTRetCMAC, "###,##0.00") & ";17; ;36;", pnItem)
'                    lnTRetCMAC = 0
'                    lnTDepCMAC = 0
'                End If
'
'                lnTRetAge = 0
'                lnTDepAge = 0
'                lbBan = True
'            End If
'
'            If Len(RegTrandiaria!cCodCta) = 18 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
'                    lbBan = True
'                End If
'            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
'                If lsGru <> RegTrandiaria!cCodAge Then
'                    lsGru = RegTrandiaria!cCodAge
'                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
'                    lbBan = True
'                End If
'            End If
'
'            If IsNull(RegTrandiaria!cNumDoc) Then
'                lsDoc = " "
'            Else
'                lsDoc = RegTrandiaria!cNumDoc
'            End If
'
'            If RegTrandiaria!nMonTran >= 0 Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
'                lnTDepCMAC = lnTDepCMAC + RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran)
'                lnTRetAge = lnTRetAge + RegTrandiaria!nMonTran
'                lnTRetCMAC = lnTRetCMAC + RegTrandiaria!nMonTran
'            End If
'
'            If IsNull(RegTrandiaria!cCodUsu) Then
'                lsCodUsu = ""
'            Else
'                lsCodUsu = RegTrandiaria!cCodUsu
'            End If
'
'            lsProd = GetProducto(RegTrandiaria!cCodCta)
'            lsOP = IIf(VerificaCtaOP(RegTrandiaria!cCodCta), "OP", "")
'
'            lsNomOpe = GetNomOpe(RegTrandiaria!cCodOpe, 40)
'            'lsDoc
'            psCadena = psCadena & Space(10 - Len(lsDoc) - 3) & Mid(gsCodAge, 4, 2) & "-" & lsDoc _
'                               & Space(2) & lsNomOpe _
'                               & Space(2) & RegTrandiaria!cCodCta _
'                               & Space(17 - Len(lsDep)) & lsDep _
'                               & Space(17 - Len(lsRet)) & lsRet _
'                               & Space(6 - Len(lsCodUsu)) & lsCodUsu _
'                               & Space(24 - Len(Format(RegTrandiaria!dFectran, "dd/mm/yyyy hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "dd/mm/yyyy hh:mm:ss AMPM") _
'                               & Space(4 - Len(lsProd)) & lsProd _
'                               & Space(3 - Len(lsOP)) & lsOP & Chr(10)
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                psCadena = psCadena & Chr(12)
'                psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, psMoneda)
'                psCadena = psCadena & Encabezado("NroDocCMAC;10;Tipo Operacion;40;Nro Cuenta;14;Deposito;17;Retiro;17;Usu;6;FechaHora;24;Pro.;4;OP;4;", pnItem)
'            End If
'
'            RegTrandiaria.MoveNext
'        Wend
'
'        If lbBan Then
'            psCadena = psCadena & Encabezado("Total Agencia;15; ;18;" & Format(lnTDepAge, "###,##0.00") & ";50;" & Format(lnTRetAge, "###,##0.00") & ";17; ;36;", pnItem)
'            psCadena = psCadena & Encabezado("Total CMAC;15; ;18;" & Format(lnTDepCMAC, "###,##0.00") & ";50;" & Format(lnTRetCMAC, "###,##0.00") & ";17; ;36;", pnItem)
'        End If
'        'psCadena = psCadena & Encabezado("Total ;30;" & Format(lnTDep, "###,##0.00") & ";53;" & Format(lnTRet, "###,##0.00") & ";17; ;36;", pnItem)
'
'        lnTRetTT = lnTRet
'        lnTDepTT = lnTDep
'
'        pnItem = pnItem + 1
'
'        RegTrandiaria.Close
'    End If
'    Set RegTrandiaria = Nothing
'
'    If lnTRetTT <> 0 Or lnTDepTT <> 0 Then
'        psCadena = psCadena & Encabezado("Total ;30;" & Format(lnTDepTT, "###,##0.00") & ";53;" & Format(lnTRetTT, "###,##0.00") & ";17; ;36;", pnItem)
'
'        psCadena = psCadena & Chr(12)
'    End If
'End Function

'Movimiento que se realiza con cuentas de esta agencia pero
'por otro servidor
Private Function MovimientoExternoAC(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String, Optional pbEsTramite As Boolean = False)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lsDep As String
    Dim lsRet As String
    Dim lsCodUsu As String
    
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnTDepTT As Double
    Dim lnTRetTT As Double
    
    Dim lbBan As Boolean
    
    Dim i As Long
    
    Dim lsGru As String
    
    Dim lnTDepAge As Double
    Dim lnTRetAge As Double
    
    Dim lsCadAdd As String
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    lnTRet = 0
    lnTDep = 0
    lnTRetTT = 0
    lnTDepTT = 0
    
    lbBan = False

    If pbEsTramite Then
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(M.cMovNro,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        End If
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
        lbBan = False
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.R;6;Hora;12;TInt;8;", pnItem)
        
        lsGru = ""
        'Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2)
        
        While Not RegTrandiaria.EOF
            lbBan = True
            If lsGru <> "" And lsGru <> RegTrandiaria!CCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            End If
            
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = RegTrandiaria!cNumDoc
            End If
            
            If RegTrandiaria!nMonTran >= 0 Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnTRetAge = lnTRetAge + RegTrandiaria!nMonTran
            End If
            
            lsCodUsu = Right(RegTrandiaria!cMovNro, 4)

            psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                                & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                                & Space(19 - Len(Format(RegTrandiaria!nSaldCnt - RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nSaldCnt - RegTrandiaria!nMonTran, "###,##0.00") _
                                & Space(19 - Len(lsDep)) & lsDep _
                                & Space(19 - Len(lsRet)) & lsRet _
                                & Space(6 - Len(lsCodUsu)) & lsCodUsu _
                                & Space(2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.R;6;Hora;12;TInt;8;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
        
        If lbBan Then psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
        
        psCadena = psCadena & Encabezado("Total ;12;" & Format(lnTDep, "###,##0.00") & ";60;" & Format(lnTRet, "###,##0.00") & ";20; ;26;", pnItem)
    
        lnTRetTT = lnTRet
        lnTDepTT = lnTDep
   
        pnItem = pnItem + 1
        
        RegTrandiaria.Close
    End If
    Set RegTrandiaria = Nothing
   
    'Ordenes de Pago
    lnTRet = 0
    lnTDep = 0
    lnTRetAge = 0
    lnTDepAge = 0
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
              & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
              & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
        If pnItem < lnRango And lbBan Then
            psCadena = psCadena & Chr(12)
        End If
        
        lbBan = False
        
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;UsuR.;6;Hora;12;TInt;8;", pnItem)
    
        lsGru = ""
        
        While Not RegTrandiaria.EOF
            
            If lsGru <> "" And lsGru <> RegTrandiaria!CCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            End If
            
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = Trim(RegTrandiaria!cNumDoc)
            End If
            
            If RegTrandiaria!nMonTran >= 0 Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnTRetAge = lnTRetAge + RegTrandiaria!nMonTran
            End If
            
            lsCodUsu = Right(RegTrandiaria!cMovNro, 4)
            
            psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                               & Space(19 - Len(Format(RegTrandiaria!nSaldCnt - RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nSaldCnt - RegTrandiaria!nMonTran, "###,##0.00") _
                               & Space(19 - Len(lsDep)) & lsDep _
                               & Space(19 - Len(lsRet)) & lsRet _
                               & Space(6 - Len(lsCodUsu)) & lsCodUsu _
                               & Space(2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & Chr(10)
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;UsuR.;6;Hora;12;TInt;8;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
        
        If lbBan Then psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
        
        psCadena = psCadena & Encabezado("Total ;12;" & Format(lnTDep, "###,##0.00") & ";60;" & Format(lnTRet, "###,##0.00") & ";20; ;26;", pnItem)
        
        lnTRetTT = lnTRetTT + lnTRet
        lnTDepTT = lnTDepTT + lnTDep
        
        pnItem = pnItem + 1
        
        RegTrandiaria.Close
    End If
    Set RegTrandiaria = Nothing
    
    If lnTRetTT <> 0 Or lnTDepTT <> 0 Then
        psCadena = psCadena & Encabezado("Total ;12;" & Format(lnTDepTT, "###,##0.00") & ";60;" & Format(lnTRetTT, "###,##0.00") & ";20; ;26;", pnItem)
        psCadena = psCadena & Chr(12)
    End If
End Function

'************************************************************************
'PLAZO FIJO
'********************************************************************************************
Private Sub ReporteMovPF(psCadenaRep As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String
    lnTotRet = 0
    lnTotDep = 0

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))
    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1

    lsTEM = psCadenaRep

    psCadenaRep = ""

    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas PlazoFijo-Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFApeEfec, , True))
    Call ReporteAperturasPF(psCadenaRep, "Movimientos en Efectivo de Cuentas de Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFMovEfec, , True))
    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas Plazo Fijo-Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFApeCheque, , True))
    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas Plazo Fijo-Transferencia" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFApeTreansferencia, , True))
    Call ReporteAperturasPF(psCadenaRep, "Movimientos Sin Efectivo de Cuentas de Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFMovSinEfec, , True))
    Call ReporteAperturasPF(psCadenaRep, "Cancelación de Cuentas Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovPFCancelacion, , True))

    If lnTotDep <> 0 Or lnTotRet <> 0 Then
        psCadenaRep = lsCebecera & psCadenaRep
        QuitarSalto psCadenaRep
        psCadenaRep = psCadenaRep & Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
        psCadenaRep = psCadenaRep & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
        psCadenaRep = psCadenaRep & Chr(12)
    End If

    lnTotRet = 0
    lnTotDep = 0

    'Call MovimientoExternosPF(psCadenaRep, "Cuentas de -Plazo Fijo- Movidas por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFRetInt & "','" & CaptacOperacion.gPFRetIntAboAho)
    'Call MovimientoExternosPF(psCadenaRep, "Cuentas Canceladas de -Plazo Fijo- por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFCancelacion)
    'Call MovimientoExternosPF(psCadenaRep, "Cuentas de -Plazo Fijo- Movidas en otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, 0, 0, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFRetInt & "','" & CaptacOperacion.gPFRetIntAboAho, True)

    psCadenaRep = lsTEM + psCadenaRep
    lsTEM = ""
End Sub

Private Sub ReporteAperturasPF(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispPF, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntPF, nPlazo Plazo, nSaldoContable nSaldCnt" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispPF, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntPF, nPlazo Plazo, nSaldoContable nSaldCnt" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & "  And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    oCon.AbreConexion
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
     If RegTrandiaria.EOF And RegTrandiaria.BOF Then
         Exit Sub
     End If
     
     psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
     psCadena = psCadena & Encabezado("Nro Cuenta;15;Nro.Doc.;20;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;9;D.Plazo;10;", pnItem)
    
     While Not RegTrandiaria.EOF
         If IsNull(RegTrandiaria!cNumDoc) Then
             lsDoc = " "
         Else
             lsDoc = RegTrandiaria!cNumDoc
         End If
         
         If RegTrandiaria!nMonTran >= 0 Then
             lnTDep = lnTDep + RegTrandiaria!nMonTran
         Else
             lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
         End If
         
         If IsNull(RegTrandiaria!nSaldCnt) Then
             lnValor = 0
         Else
             lnValor = RegTrandiaria!nSaldCnt
         End If
         
         psCadena = psCadena & Space(19 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                             & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                             & Space(18 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                             & Space(9 - Len(Right(RegTrandiaria!cMovNro, 4))) & Right(RegTrandiaria!cMovNro, 4) _
                             & Space(15 - 6) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                             & Space(20 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                             & Space(9 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
                             & Space(10 - Len(Str(RegTrandiaria!Plazo))) & Str(RegTrandiaria!Plazo) & Chr(10)
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & Chr(12)
             psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
             psCadena = psCadena & Encabezado("Nro Cuenta;15;Nro.Doc.;20;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;9;D.Plazo;10", pnItem)
         End If
         
         RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15;" & Trim(Format(lnTDep, "###,##0.00")) & ";40; ;63;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub

'Operacioens hechas por otras agencias en este servidor
Private Function MovimientoExternosPF(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String = "", Optional pbEsTramite As Boolean = False)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lsDep As String
    Dim lsRet As String
    
    Dim lnTDep As Double
    Dim lnTRet As Double
    
    Dim lnTDepT As Double
    Dim lnTRetT As Double
    Dim lnValor As Double
    Dim lnOpe As Double
    
    Dim lsGru As String
    
    Dim lbBan As Boolean
    
    Dim lnTRetAge As Double
    Dim lnTDepAge As Double
    Dim lsCodUsuX As String
    
    
    lnTRetAge = 0
    lnTDepAge = 0

    
    Dim i As Long
    
    lnTRet = 0
    lnTDep = 0
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If pbEsTramite Then
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispPF, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(M.cMovNro,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispPF, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge , nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        End If
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.BOF And RegTrandiaria.EOF) Then
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;IntDevAnt;20;Retiro;20;Usu.;10;Hora;17;TInt;9;Plazo;10;", pnItem)
        
        lsGru = ""
        lbBan = False
        
        While Not RegTrandiaria.EOF
            lbBan = True
            If lsGru <> "" And lsGru <> RegTrandiaria!CCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15; ;37;" & Format(lnTRetAge * -1, "###,##0.00") & ";20; ;46;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            
'
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            End If

            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = Trim(RegTrandiaria!cNumDoc)
            End If
            
            If RegTrandiaria!nMonTran >= 0 Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnOpe = RegTrandiaria!nMonTran
                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnOpe = RegTrandiaria!nMonTran
                lnTRetAge = lnTRetAge + RegTrandiaria!nMonTran
            End If
            
            If IsNull(RegTrandiaria!nIntDev) Then
                lnValor = 0
            Else
                lnValor = RegTrandiaria!nIntDev
            End If
            
            If IsNull(RegTrandiaria!cCodUsu) Then
                If IsNull(RegTrandiaria!cCodusurem) Then
                    lsCodUsuX = ""
                Else
                    lsCodUsuX = RegTrandiaria!cCodusurem
                End If
            Else
                lsCodUsuX = RegTrandiaria!cCodUsu
            End If
            
            psCadena = psCadena & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                               & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                               & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
                               & Space(20 - Len(lsRet)) & lsRet _
                               & Space(10 - Len(lsCodUsuX)) & lsCodUsuX _
                               & Space(17 - Len(Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM") _
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
                               & Space(11 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & Chr(10)
            
            'Str(RegTrandiaria!nPlazo)
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;IntDevAnt;20;Retiro;20;Usu.;10;Hora;17;TInt;9;Plazo;10;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
        
        If lbBan Then psCadena = psCadena & Encabezado("Total Agencia;15; ;37;" & Format(lnTRetAge * -1, "###,##0.00") & ";20; ;46;", pnItem)
        
        lnTRetT = lnTRet
        lnTDepT = lnTDep
            
        RegTrandiaria.Close
    End If
    
    Set RegTrandiaria = Nothing
    
    If lnTRetT <> 0 Or lnTDepT <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15; ;37;" & Format(lnTRetT, "###,##0.00") & ";20; ;46;", pnItem)
        pnTotRet = pnTotRet + lnTRetT
        pnTotDep = pnTotDep + lnTDepT
        
        psCadena = psCadena & Chr(12)
    End If
End Function


'************************************************************************
'CTS
'********************************************************************************************
Private Sub ReporteMovCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")

    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))

    lnTotRet = 0
    lnTotDep = 0

    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1
    lsTEM = psCadena

    psCadena = ""

    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSApeEfec, , True))
    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS con Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSApeCheque, , True))
    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS con Trnasferencia" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSApeTreansferencia, , True))
    Call ReporteAperturasCTS(psCadena, "Movimientos en Efectivo de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSMovEfec, , True))
    Call ReporteAperturasCTS(psCadena, "Movimientos Sin Efectivo de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSMovSinEfec, , True))
    Call ReporteAperturasCTS(psCadena, "Cancelación de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSCancelacion, , True))
    Call ReporteAperturasCTS(psCadena, "Depositos con Cheque en Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, GetOpeGrupoRep(gCtasMovCTSDepCheque, , True))

    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadena = lsCebecera + psCadena
        QuitarSalto psCadena
        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
        psCadena = psCadena & Chr(12)
    End If

    psCadena = lsTEM + psCadena
    lsTEM = ""
End Sub

'************************************************************************
'OPeraciones de Creditos
'********************************************************************************************
Private Sub ReporteMovCreditos(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")

    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))

    lnTotRet = 0
    lnTotDep = 0

    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1
    lsTEM = psCadena

    psCadena = ""
    lnRango = 58
    
    Call ReporteMovimientosCreditos(psCadena, "Operaciones de Creditos" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadena = lsCebecera + psCadena
        QuitarSalto psCadena
        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(Abs(lnTotRet), "###,##0.00")) & ";23;" & Trim(Format(lnTotDep, "###,##0.00")) & ";15; ;26;", pnItem)
        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";37; ;46;", pnItem)
        psCadena = psCadena & Chr(12)
    End If

    psCadena = lsTEM + psCadena
    lsTEM = ""
End Sub


Public Sub ReporteMovimientosCreditos(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, CASE WHEN MC.cOpeCod like '1001__' THEN "
        Sql1 = Sql1 & "        (Select SUM(nMonto) From MovColDet where nMovNro = MC.nMovNro and cOpeCod like '1001__')"
        Sql1 = Sql1 & "          END as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN MC.cOpeCod like '100[234567]__' OR or MC.cOpeCod in ('107001') THEN"
        Sql1 = Sql1 & " (Select SUM(nMonto) From MovColDet where nMovNro = MC.nMovNro and ( cOpeCod like '100[234567]__' OR cOpeCod in ('107001') ))"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where cOpeCod like '100[1234567]__' or cOpeCod in ('107001')"
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%' ORDER BY MC.cCtaCod"
            
    Else
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, CASE WHEN MC.cOpeCod like '1001__' THEN "
        Sql1 = Sql1 & "        (Select SUM(nMonto) From MovColDet where nMovNro = MC.nMovNro and cOpeCod like '1001__')"
        Sql1 = Sql1 & "          END as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN MC.cOpeCod like '100[234567]__' OR MC.cOpeCod in ('107001') THEN"
        Sql1 = Sql1 & " (Select SUM(nMonto) From MovColDet where nMovNro = MC.nMovNro and ( cOpeCod like '100[234567]__' OR cOpeCod in ('107001') ))"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where cOpeCod like '100[1234567]__' OR cOpeCod in ('107001') "
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY MC.cCtaCod"
            
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
        Exit Sub
    End If

    psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
          
    While Not RegTrandiaria.EOF
                
        lsDoc = " "
                
        lnTRet = lnTRet - IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso)
        lnTDep = lnTDep + IIf(IsNull(RegTrandiaria!nPagos), 0, Abs(RegTrandiaria!nPagos))
                
        lnValor = 0
                
        psCadena = psCadena & Space(23 - Len(RegTrandiaria!Cctacod)) & RegTrandiaria!Cctacod _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00") _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cUser)) & RegTrandiaria!cUser _
                    & Space(17 - Len(Format(RegTrandiaria!Hora, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!Hora, "hh:mm:ss AMPM") _
                    & Chr(10)
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & Chr(12)
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        'psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub

'************************************************************************
'OPeraciones de Pignoraticio
'********************************************************************************************
Private Sub ReporteMovPignoraticio(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")

    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))

    lnTotRet = 0
    lnTotDep = 0

    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1
    lsTEM = psCadena

    psCadena = ""
    lnRango = 58
    Call ReporteMovimientosPignoraticio(psCadena, "Operaciones de Pignoraticio" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadena = lsCebecera + psCadena
        QuitarSalto psCadena
        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(Abs(lnTotRet), "###,##0.00")) & ";23;" & Trim(Format(lnTotDep, "###,##0.00")) & ";15; ;26;", pnItem)
        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";37; ;46;", pnItem)
        psCadena = psCadena & Chr(12)
    End If

    psCadena = lsTEM + psCadena
    lsTEM = ""
End Sub


Public Sub ReporteMovimientosPignoraticio(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, CASE WHEN MC.cOpeCod = '120201' THEN "
        Sql1 = Sql1 & "        (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & "          END as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN SUBSTRING(MC.cOpeCod,1,4) in ('1210','1211','1212','1216','1219','1261','1262','1263') THEN"
        Sql1 = Sql1 & " (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where SUBSTRING(cOpeCod,1,4) in ('1210','1211','1212','1216','1219','1261','1262','1263')"
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%' ORDER BY MC.cCtaCod"
            
    Else
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, CASE WHEN MC.cOpeCod = '120201' THEN "
        Sql1 = Sql1 & "        (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & "          END as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN SUBSTRING(MC.cOpeCod,1,4) in ('1210','1211','1212','1216','1219','1261','1262','1263') THEN"
        Sql1 = Sql1 & "  (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where SUBSTRING(cOpeCod,1,4) in ('1210','1211','1212','1216','1219','1261','1262','1263')"
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY MC.cCtaCod"
            
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
        Exit Sub
    End If

    psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
          
    While Not RegTrandiaria.EOF
                
        lsDoc = " "
                
        lnTRet = lnTRet - IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso)
        lnTDep = lnTDep + IIf(IsNull(RegTrandiaria!nPagos), 0, Abs(RegTrandiaria!nPagos))
                
        lnValor = 0
                
        psCadena = psCadena & Space(23 - Len(RegTrandiaria!Cctacod)) & RegTrandiaria!Cctacod _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00") _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cUser)) & RegTrandiaria!cUser _
                    & Space(17 - Len(Format(RegTrandiaria!Hora, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!Hora, "hh:mm:ss AMPM") _
                    & Chr(10)
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & Chr(12)
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        'psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub


'************************************************************************
'OPeraciones de Judicial
'********************************************************************************************
Private Sub ReporteMovJudicial(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")

    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))

    lnTotRet = 0
    lnTotDep = 0

    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1
    lsTEM = psCadena

    psCadena = ""
    lnRango = 58
    Call ReporteMovimientosJudicial(psCadena, "Operaciones de Judicial" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadena = lsCebecera + psCadena
        QuitarSalto psCadena
        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(Abs(lnTotRet), "###,##0.00")) & ";23;" & Trim(Format(lnTotDep, "###,##0.00")) & ";15; ;26;", pnItem)
        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";37; ;46;", pnItem)
        psCadena = psCadena & Chr(12)
    End If

    psCadena = lsTEM + psCadena
    lsTEM = ""
End Sub


Public Sub ReporteMovimientosJudicial(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, 0 as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN SUBSTRING(MC.cOpeCod,1,4) in ('1302','1303','1304','1306') THEN"
        Sql1 = Sql1 & " (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where SUBSTRING(cOpeCod,1,4) in ('1302','1303','1304','1306')"
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%' ORDER BY MC.cCtaCod"
            
    Else
    
        Sql1 = " Select M.cMovNro, MC.cCtaCod, 0 as nDesembolso,"
        Sql1 = Sql1 & " CASE WHEN SUBSTRING(MC.cOpeCod,1,4) in ('1302','1303','1304','1306') THEN"
        Sql1 = Sql1 & " (Select SUM(nMonto) From MovCol where nMovNro = MC.nMovNro)"
        Sql1 = Sql1 & " END As nPagos, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " From Mov M"
        Sql1 = Sql1 & " Inner Join ("
        Sql1 = Sql1 & " Select * from MovCol where SUBSTRING(cOpeCod,1,4) in ('1302','1303','1304','1306')"
        Sql1 = Sql1 & " ) MC ON M.nMovNro = MC.nMovNro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 and  Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' "
        Sql1 = Sql1 & sqlAdd
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY MC.cCtaCod"
            
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
        Exit Sub
    End If

    psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
          
    While Not RegTrandiaria.EOF
                
        lsDoc = " "
                
        lnTRet = lnTRet - IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso)
        lnTDep = lnTDep + IIf(IsNull(RegTrandiaria!nPagos), 0, Abs(RegTrandiaria!nPagos))
                
        lnValor = 0
                
        psCadena = psCadena & Space(23 - Len(RegTrandiaria!Cctacod)) & RegTrandiaria!Cctacod _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nDesembolso), 0, RegTrandiaria!nDesembolso), "###,##0.00") _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nPagos), 0, RegTrandiaria!nPagos), "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cUser)) & RegTrandiaria!cUser _
                    & Space(17 - Len(Format(RegTrandiaria!Hora, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!Hora, "hh:mm:ss AMPM") _
                    & Chr(10)
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & Chr(12)
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Nro Cuenta;20;Desembolso;20;Pagos;12;Usu.;10;Hora;15;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        'psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub


'************************************************************************
'OPeraciones de CompraVenta
'********************************************************************************************
Private Sub ReporteCompraVenta(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")

    Dim Sql1 As String
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lnDep As Double
    Dim lnRet As Double

    Dim lnTotDep As Double
    Dim lnTotRet As Double
    Dim lsCebecera As String

    Dim lsTEM As String
    Dim i As Long
    Dim lsAdd As String

    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))

    lnTotRet = 0
    lnTotDep = 0

    If psCodUser = "XXX" Then
        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If

    pnPagina = pnPagina - 1
    lsTEM = psCadena

    psCadena = ""
    lnRango = 58
    Call ReporteMovimientosCompraVenta(psCadena, "Operaciones de Compra Venta" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadena = lsCebecera + psCadena
        QuitarSalto psCadena
        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";23;" & Trim(Format(lnTotRet, "###,##0.00")) & ";15; ;26;", pnItem)
        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";37; ;46;", pnItem)
        psCadena = psCadena & Chr(12)
    End If

    psCadena = lsTEM + psCadena
    lsTEM = ""
End Sub


Public Sub ReporteMovimientosCompraVenta(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "")
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
    
        Sql1 = " select O.cOpeDesc, MV.nMovImporte, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " from mov M Inner Join Opetpo O ON M.cOpeCod = O.cOpeCod"
        Sql1 = Sql1 & " Inner Join MovCompraVenta MV ON MV.nMovnro = M.nMovnro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 "
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%' ORDER BY M.nMovNro"
            
    Else
    
        Sql1 = " select O.cOpeDesc, MV.nMovImporte, RIGHT(m.cMovNro, 4) As cUser, SUBSTRING(m.cMovNro, 9, 2) + ':' + SUBSTRING(m.cMovNro, 11, 2) + ':' + SUBSTRING(m.cMovNro, 13, 2) As Hora"
        Sql1 = Sql1 & " from mov M Inner Join Opetpo O ON M.cOpeCod = O.cOpeCod"
        Sql1 = Sql1 & " Inner Join MovCompraVenta MV ON MV.nMovnro = M.nMovnro"
        Sql1 = Sql1 & " Where M.nMovFlag =0 "
        Sql1 = Sql1 & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY M.nMovNro"
            
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
        Exit Sub
    End If

    psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    psCadena = psCadena & Encabezado("Operacion;20;Monto;20;Usu.;10;Hora;15;", pnItem)
          
    While Not RegTrandiaria.EOF
                
        lsDoc = " "
                
        lnTRet = lnTRet + IIf(IsNull(RegTrandiaria!nMovImporte), 0, RegTrandiaria!nMovImporte)
        lnTDep = lnTDep + IIf(IsNull(RegTrandiaria!nMovImporte), 0, Abs(RegTrandiaria!nMovImporte))
                
        lnValor = 0
                
        psCadena = psCadena & Space(23 - Len(RegTrandiaria!cOpeDesc)) & RegTrandiaria!cOpeDesc _
                    & Space(15 - Len(Format(IIf(IsNull(RegTrandiaria!nMovImporte), 0, RegTrandiaria!nMovImporte), "###,##0.00"))) & Format(IIf(IsNull(RegTrandiaria!nMovImporte), 0, RegTrandiaria!nMovImporte), "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cUser)) & RegTrandiaria!cUser _
                    & Space(17 - Len(Format(RegTrandiaria!Hora, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!Hora, "hh:mm:ss AMPM") _
                    & Chr(10)
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & Chr(12)
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Operacion;30;Monto;20;Usu.;10;Hora;15;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        'psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub


Public Sub ReporteAperturasCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim i As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2) = '" & lsAgeCod & "' "
    End If
    
    If psCodUser = "XXX" Then
        Sql1 = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, P.nSaldo nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaInt, Right(cMovNro,4) cCodUsu" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC , P.nSaldo nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaInt, Right(cMovNro,4) cCodUsu " _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) " & sqlAdd & " And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If RegTrandiaria.EOF And RegTrandiaria.BOF Then
        Exit Sub
    End If

    psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usu.;10;Hora;17;Sald.Cnt;22;TInt;11;", pnItem)
   
    lnPorRet = 0.5 '(ReadParametros("23110") / 100)
   
    While Not RegTrandiaria.EOF
        
        If IsNull(RegTrandiaria!cNumDoc) Then
            lsDoc = " "
        Else
            lsDoc = RegTrandiaria!cNumDoc
        End If
        
        If RegTrandiaria!nMonTran >= 0 And Left(RegTrandiaria!copecod, 4) <> "2203" Then
            lnTDep = lnTDep + RegTrandiaria!nMonTran
        Else
            lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
        End If
        
        If IsNull(RegTrandiaria!nSaldCnt) Then
            lnValor = 0
        Else
            lnValor = RegTrandiaria!nSaldCnt
        End If
        
        psCadena = psCadena & Space(18 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                    & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                    & Space(22 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
                    & Space(17 - Len(Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM") _
                    & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                    & Space(11 - Len(Format(RegTrandiaria!nTasaInt, "###,##0.00"))) & Format(RegTrandiaria!nTasaInt, "###,##0.00") & Chr(10)
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & Chr(12)
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usu.;10;Hora;17;Sald.Cnt;22;TInt;11;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & Chr(12)
    End If
End Sub

'Operacioens hechas por otras agencias en este servidor
Private Sub MovimientoExternosCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As Moneda = gMonedaNacional, Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String = "", Optional pbEsTramite As Boolean = False)
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim lsDep As String
    Dim lsRet As String
    
    Dim lnTDep As Double
    Dim lnTRet As Double
    
    Dim lnTDepT As Double
    Dim lnTRetT As Double
    Dim lnValor As Double
    Dim lnOpe As Double
    Dim lnPorRet As Double
    Dim lbBan  As Boolean
    Dim lsGru  As String
    Dim lnTRetAge As Currency
    Dim lnTDepAge As Currency
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim i As Long
    
    lnTRet = 0
    lnTDep = 0
    
    If pbEsTramite Then
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                 & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge" _
                 & " From Mov M" _
                 & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                 & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                 & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                 & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                 & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                 & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                 & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                 & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                 & " And Substring(M.cMovNro,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,18,2) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) Like '___" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        End If
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.BOF And RegTrandiaria.EOF) Then
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
        
        lsGru = ""
        lnTRetAge = 0
        lnTDepAge = 0
        
        While Not RegTrandiaria.EOF
            lbBan = True
            If lsGru <> "" And lsGru <> RegTrandiaria!CCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;46;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If

            If Len(RegTrandiaria!cCodCta) = 12 Then
                If lsGru <> RegTrandiaria!CCodAge Then
                    lsGru = RegTrandiaria!CCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            End If
           
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = RegTrandiaria!cNumDoc
            End If
            
            If RegTrandiaria!nMonTran >= 0 Then
                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
                lsRet = "0.00"
                lnTDep = lnTDep + RegTrandiaria!nMonTran
                lnOpe = RegTrandiaria!nMonTran
                lnTDepAge = lnTDepAge + RegTrandiaria!nMonTran
            Else
                lsDep = "0.00"
                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
                lnOpe = RegTrandiaria!nMonTran
                lnTRetAge = lnTRetAge + (RegTrandiaria!nMonTran * -1)
            End If
            
            If IsNull(RegTrandiaria!nSaldCnt) Then
                lnValor = 0
            Else
                lnValor = RegTrandiaria!nSaldCnt
            End If
            
            psCadena = psCadena & Space(14 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                                & Space(18 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                                & Space(20 - Len(Format(lnValor - lnOpe, "###,##0.00"))) & Format(lnValor - lnOpe, "###,##0.00") _
                                & Space(20 - Len(lsDep)) & lsDep _
                                & Space(20 - Len(lsRet)) & lsRet _
                                & Space(6 - Len(RegTrandiaria!cCodusurem)) & RegTrandiaria!cCodusurem _
                                & Space(12 - Len(Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFecTran, "hh:mm:ss AMPM") _
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & Chr(10)
                               
            pnItem = pnItem + 1
            
            If pnItem = lnRango Then
                psCadena = psCadena & Chr(12)
                psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;Sald.Cnt.;20;TInt;8;", pnItem)
            End If
            
            RegTrandiaria.MoveNext
        Wend
            
        If lbBan Then psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;46;", pnItem)
        lnTRetT = lnTRet
        lnTDepT = lnTDep
            
        RegTrandiaria.Close
    End If
        
    Set RegTrandiaria = Nothing
    
    If lnTRetT <> 0 Or lnTDepT <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15;" & Format(lnTDepT, "###,##0.00") & ";57;" & Format(lnTRetT, "###,##0.00") & ";20; ;25;", pnItem)
        pnTotRet = pnTotRet + lnTRetT
        pnTotDep = pnTotDep + lnTDepT
        psCadena = psCadena & Chr(12)
    End If
End Sub

Public Function SdosTipCta(Optional psCodAge As String = "") As String
    Dim sPers As String * 31
    Dim lnNumRec As Currency
    Dim lnSdoPro As Currency
    Dim lnSdoTot As Currency
    Dim lnRecPro As Currency
    Dim sqlAgencia As String
    
    Dim lsMoneda As String
    Dim m As Moneda
    Dim P As PersPersoneria
    Dim lsQrySdo As String
    Dim rsSdo As New ADODB.Recordset
    Dim lnRegCod As Integer
    Dim lsValor As String, lsNumPag As String
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsTipMon As String, lsTipPer As String
    Dim lnCarLin As Integer, lnCntPag As Integer
    Dim lnRecs01 As Integer, lnRecs02 As Integer
    Dim lnRecs03 As Integer, lnRecs04 As Integer
    Dim lnRecs05 As Integer, lnRecs06 As Integer
    Dim lnSdos01 As Currency, lnSdos02 As Currency
    Dim lnSdos03 As Currency, lnSdos04 As Currency
    Dim lnSdos05 As Currency, lnSdos06 As Currency
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim sqlAdd As String
    
    lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0: lnRecs05 = 0: lnRecs06 = 0
    lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0: lnSdos05 = 0: lnSdos06 = 0
    
    lnCarLin = 170
    
    oCon.AbreConexion
    
    Set rsSdo = New ADODB.Recordset
    rsSdo.CursorLocation = adUseClient
    
    If psCodAge = "" Then
        sqlAgencia = ""
    Else
        sqlAgencia = " And Left(PR.cCtaCod,5) Like  '___" & psCodAge & "'"
    End If
    
    lsRTFRfm = ""
    
    For m = gMonedaNacional To gMonedaExtranjera
        lnCntPag = 1
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        lsMoneda = IIf(m = Moneda.gMonedaNacional, "SOLES", "DOLARES")
        lsTitRp1 = "I N F O R M E   D E   S A L D O S   P O R   T I P O   D E   C U E N T A"
        lsTitRp2 = ""
        lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(38, " ") & "        A  H  O  R  R  O  S               P L A Z O    F I J O                     C    T    S                       T  O  T  A  L  " & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(38, " ") & "A C T I V A S        I N A C T I V A S" & String(29, " ") & "C L I E N T E S        E M P L E A D O S      A C U M U L A D O" & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(8, " ") & "C O N C E P T O S" & String(8, " ")
        lsRTFRfm = lsRTFRfm & "NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S" & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        
        For P = gPersonaNat To gPersonaJurCFLEdpyme
            If (P < gPersonaJurCFLCMAC Or P >= gPersonaJurCFLCMAC) Then
                Select Case P
                    Case PersPersoneria.gPersonaNat
                        sPers = "PERSONA NATURAL"
                    Case PersPersoneria.gPersonaJurSFL
                        sPers = "PERSONA JUR. SIN FINES DE LUCRO"
                    Case PersPersoneria.gPersonaJurCFL
                        sPers = "PERSONA JUR. CON FINES DE LUCRO"
                    Case PersPersoneria.gPersonaJurCFLCMAC
                        sPers = "CAJAS MUNICIPALES"
                    Case PersPersoneria.gPersonaJurCFLCRAC
                        sPers = "CAJAS RURALES"
                    Case PersPersoneria.gPersonaJurCFLCooperativa
                        sPers = "COOPERATIVAS"
                    Case PersPersoneria.gPersonaJurCFLEdpyme
                        sPers = "EDPYMES"
                End Select
                lsRTFRfm = lsRTFRfm & sPers & "  "
                
                lsQrySdo = " SELECT ISNULL(COUNT(PR.cCtaCod),0) Num, ISNULL(SUM(nSaldo),0) nSdoCnt FROM Producto PR" _
                         & " Inner Join Captaciones CA On PR.cCtaCod = CA.cCtaCod Inner Join CaptacAhorros CAA On PR.cCtaCod = CAA.cCtaCod WHERE Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND bInactiva = 0" _
                         & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & m & "' AND nPersoneria = '" & P & "'" & sqlAgencia
                If rsSdo.State = adStateOpen Then rsSdo.Close
                Set rsSdo = oCon.CargaRecordSet(lsQrySdo)
                If rsSdo.EOF And rsSdo.BOF Then
                    lnNumRec = 0: lnSdoPro = 0
                Else
                    lnNumRec = rsSdo!Num
                    lnSdoPro = rsSdo!nSdoCnt
                    lnSdoTot = lnSdoTot + lnSdoPro
                    lnRecPro = lnRecPro + lnNumRec
                End If
          
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnNumRec)), 6, False, 6, 0) & " "
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoPro)), 15, True, 9, 2) & " "
                lnRecs01 = lnRecs01 + lnNumRec
                lnSdos01 = lnSdos01 + lnSdoPro
        
                lsQrySdo = " SELECT ISNULL(COUNT(PR.cCtaCod),0) Num, ISNULL(SUM(nSaldo),0) nSdoCnt FROM Producto PR" _
                         & " Inner Join Captaciones CA On PR.cCtaCod = CA.cCtaCod Inner Join CaptacAhorros CAA On PR.cCtaCod = CAA.cCtaCod WHERE Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND bInactiva = 1" _
                         & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & m & "' AND nPersoneria = '" & P & "'" & sqlAgencia
                If rsSdo.State = adStateOpen Then rsSdo.Close
                Set rsSdo = oCon.CargaRecordSet(lsQrySdo)
                
                If rsSdo.EOF And rsSdo.BOF Then
                    lnNumRec = 0: lnSdoPro = 0
                Else
                    lnNumRec = rsSdo!Num
                    lnSdoPro = rsSdo!nSdoCnt
                    lnSdoTot = lnSdoTot + lnSdoPro
                    lnRecPro = lnRecPro + lnNumRec
                End If
        
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnNumRec)), 6, False, 6, 0) & " "
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoPro)), 15, True, 9, 2) & " "
                lnRecs02 = lnRecs02 + lnNumRec
                lnSdos02 = lnSdos02 + lnSdoPro
                
                lsQrySdo = " SELECT ISNULL(COUNT(PR.cCtaCod),0) Num, ISNULL(SUM(nSaldo),0) nSdoCnt FROM Producto PR" _
                         & " Inner Join Captaciones CA On PR.cCtaCod = CA.cCtaCod Inner Join CaptacPlazoFijo CAA On PR.cCtaCod = CAA.cCtaCod WHERE Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                         & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & m & "' AND nPersoneria = '" & P & "'" & sqlAgencia
                
                If rsSdo.State = adStateOpen Then rsSdo.Close
                Set rsSdo = oCon.CargaRecordSet(lsQrySdo)

                Set rsSdo.ActiveConnection = Nothing
                If rsSdo.EOF And rsSdo.BOF Then
                    lnNumRec = 0: lnSdoPro = 0
                Else
                    lnNumRec = rsSdo!Num
                    lnSdoPro = rsSdo!nSdoCnt
                    lnSdoTot = lnSdoTot + lnSdoPro
                    lnRecPro = lnRecPro + lnNumRec
                End If
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnNumRec)), 6, False, 6, 0) & " "
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoPro)), 15, True, 9, 2) & " "
                lnRecs03 = lnRecs03 + lnNumRec
                lnSdos03 = lnSdos03 + lnSdoPro
                    
                If P <> gPersonaJurCFLCMAC Then
                              
                    lsQrySdo = " SELECT ISNULL(COUNT(PR.cCtaCod),0) Num, ISNULL(SUM(nSaldo),0) nSdoCnt FROM Producto PR" _
                             & " Inner Join Captaciones CA On PR.cCtaCod = CA.cCtaCod Inner Join CaptacCTS CAA On PR.cCtaCod = CAA.cCtaCod WHERE Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                             & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & m & "' AND nPersoneria = '" & P & "' AND cCodInst <> '" & gsInstCmac & "'" & sqlAgencia
                    
                    If rsSdo.State = adStateOpen Then rsSdo.Close
                    Set rsSdo = oCon.CargaRecordSet(lsQrySdo)
                    
                    If rsSdo.EOF And rsSdo.BOF Then
                       lnNumRec = 0: lnSdoPro = 0
                    Else
                       lnNumRec = rsSdo!Num
                       lnSdoPro = rsSdo!nSdoCnt
                       lnSdoTot = lnSdoTot + lnSdoPro
                       lnRecPro = lnRecPro + lnNumRec
                    End If
        
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnNumRec)), 6, False, 6, 0) & " "
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoPro)), 15, True, 9, 2) & " "
                    lnRecs04 = lnRecs04 + lnNumRec
                    lnSdos04 = lnSdos04 + lnSdoPro
                  
                    lsQrySdo = " SELECT ISNULL(COUNT(PR.cCtaCod),0) Num, ISNULL(SUM(nSaldo),0) nSdoCnt FROM Producto PR" _
                             & " Inner Join Captaciones CA On PR.cCtaCod = CA.cCtaCod Inner Join CaptacCTS CAA On PR.cCtaCod = CAA.cCtaCod WHERE Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                             & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & m & "' AND nPersoneria = '" & P & "' AND cCodInst = '" & gsInstCmac & "'" & sqlAgencia
                    
                    If rsSdo.State = adStateOpen Then rsSdo.Close
                    Set rsSdo = oCon.CargaRecordSet(lsQrySdo)
                    
                    Set rsSdo.ActiveConnection = Nothing
                    If rsSdo.EOF And rsSdo.BOF Then
                        lnNumRec = 0: lnSdoPro = 0
                    Else
                        lnNumRec = rsSdo!Num
                        lnSdoPro = rsSdo!nSdoCnt
                        lnSdoTot = lnSdoTot + lnSdoPro
                        lnRecPro = lnRecPro + lnNumRec
                    End If
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnNumRec)), 6, False, 6, 0) & " "
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoPro)), 15, True, 9, 2) & " "
                    lnRecs05 = lnRecs05 + lnNumRec
                    lnSdos05 = lnSdos05 + lnSdoPro
                Else
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(0)), 6, False, 6, 0) & " "
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(0)), 15, True, 9, 2) & " "
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(0)), 6, False, 6, 0) & " "
                    lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(0)), 15, True, 9, 2) & " "
                End If
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecPro)), 6, False, 6, 0) & " "
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoTot)), 15, True, 9, 2) & Chr$(10)
                lnRecs06 = lnRecs06 + lnRecPro
                lnSdos06 = lnSdos06 + lnSdoTot
                  
                lnSdoTot = 0: lnRecPro = 0
            End If
           
        Next P
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        lsRTFRfm = lsRTFRfm & "T  O  T  A  L  E  S" & String(14, " ")
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs01)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos01)), 15, True, 9, 2) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs02)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos02)), 15, True, 9, 2) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs03)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos03)), 15, True, 9, 2) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs04)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos04)), 15, True, 9, 2) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs05)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos05)), 15, True, 9, 2) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs06)), 6, False, 6, 0) & " "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos06)), 15, True, 9, 2) & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(10)
        If m < 2 Then
            lsRTFRfm = lsRTFRfm & Chr$(12)
            lsRTFRfm = lsRTFRfm & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
            lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0: lnRecs05 = 0: lnRecs06 = 0
            lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0: lnSdos05 = 0: lnSdos06 = 0
        End If
    Next m
    SdosTipCta = lsRTFRfm
End Function


'   ------------------------------------------------------------
'   Función _GlobalNameSpace
'   Propósito   :   Define cabecera del reporte
'   Parámetro(s):   pnCarLin -> Caracteres x linea del reporte
'                   psSeccio -> Area responsable
'                   psTitRp1 -> Titulo 1 del reporte
'                   psTitRp2 -> Titulo 2 del reporte
'                   psMoneda -> Tipo de Moneda
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Formato     :   CabeRepo()
Private Function CabeRepo(psCabe01 As String, psCabe02 As String, _
                         pnCarLin As Integer, psSeccio As String, _
                         psTitRp1 As String, psTitRp2 As String, _
                         psMoneda As String, psNumPag As String) As String

    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsMoneda As String
    
    lsTitRp1 = "": lsTitRp2 = ""
    CabeRepo = ""
    lsMoneda = ""
    lsMoneda = IIf(psMoneda = "", String(10, " "), " - " & psMoneda)
'    psCabe01 = ""
'    psCabe02 = ""
    
    '   Definición de Cabecera 1
 
    psCabe01 = FillText(UCase(Trim(lsNomAge)) & lsMoneda, 36, " ")
    psCabe01 = psCabe01 & Space((pnCarLin - 36) - (Len(psCabe01) - 2))
    psCabe01 = psCabe01 & "PAGINA: " & psNumPag
    psCabe01 = psCabe01 & Space(5) & "FECHA: " & Format(ldFecRepo, gsFormatoFechaView)
    '   Definición de Cabecera 2
    psCabe02 = FillText(psSeccio, 19, " ")
    psCabe02 = psCabe02 & Space((pnCarLin - 19) - (Len(psCabe02) - 2))
    psCabe02 = psCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
    lsTitRp1 = String(Int((pnCarLin - Len(psTitRp1)) / 2), " ") & psTitRp1
    lsTitRp2 = String(Int((pnCarLin - Len(psTitRp2)) / 2), " ") & psTitRp2
    
    CabeRepo = CabeRepo & psCabe01 & Chr$(10)
    CabeRepo = CabeRepo & psCabe02 & Chr$(10)
    CabeRepo = CabeRepo & lsTitRp1 & Chr$(10)
    CabeRepo = CabeRepo & lsTitRp2
        
End Function


'   ------------------------------------------------------------
'   Función     :   JDNum
'   Propósito   :   Justifica a la derecha un campo numérico en
'                   un reporte
'   Uso         :   Reportes en general
'   Parámetro(s):   pnCampos -> Importe
'                   pnLongit -> Longitud del Campo
'                   pbComass -> Número con separador de miles
'                   pnDigEnt -> Cantidad de digitos enteros
'                   pnDigDec -> Cantidad de digitos decimales
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'
Private Function JDNum(pnCampos As String, pnLongit As Integer, _
                      pbComass As Boolean, pnDigEnt As Integer, _
                      pnDigDec As Integer)
                      
    Dim Formato As String, i As Integer, lnPosDig As Integer
    
    If pnCampos = "0.00" Then
       JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
       Exit Function
    End If
    
    If pbComass Then
       lnPosDig = 0
       For i = 1 To pnDigEnt
           lnPosDig = lnPosDig + 1
           Select Case lnPosDig
              Case 1
                   Formato = "0" & Formato
              Case 4, 7, 10
                   Formato = "#," & Formato
              Case Else
                   Formato = "#" & Formato
           End Select
       Next i
       If pnDigDec > 0 Then
          Formato = Formato & "." & String(pnDigDec, "0")
       End If
    Else
       For i = 1 To pnDigEnt
           Formato = IIf(i = 1, "0", "#") & Formato
       Next i
       If pnDigDec > 0 Then
          Formato = Formato & "." & String(pnDigDec, "0")
       End If
    End If
    pnCampos = Format(pnCampos, Formato)
    
    JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
    
End Function


Public Function EstratDep(Optional psCodAge As String = "") As String
    Dim m As Moneda, P As Integer
    Dim lsQrySdo As String
    Dim rsQrySdo As New ADODB.Recordset
    Dim lnRegCod As Integer
    Dim lsValor As String, lsNumPag As String
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsTipMon As String, lsTipPer As String
    Dim lnCarLin As Integer, lnCntPag As Integer
    Dim lnRecs01 As Integer, lnRecs02 As Integer
    Dim lnRecs03 As Integer, lnRecs04 As Integer
    Dim lnSdos01 As Currency, lnSdos02 As Currency
    Dim lnSdos03 As Currency, lnSdos04 As Currency
    Dim lnSaldos As Currency
    Dim lsLimite As String, lsLimInf As String, lsLimSup As String
    Dim i As Integer, J As Integer
    Dim lnTCFHoy As Currency, lnMinVal As Currency, lnMaxVal As Currency
    Dim lnLimInd As Integer
  
    Dim lsMoneda As String
    Dim lnRecPro As Currency
    Dim lnSdoTot As Currency
    Dim sqlAgencia As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lsRTFRfm = ""
    
    If psCodAge = "" Then
        sqlAgencia = ""
    Else
        sqlAgencia = " And Left(PR.cCtaCod,5) Like '___" & psCodAge & "'"
    End If
    
    lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0
    lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0
    
    lnCarLin = 130
    lsLimite = "     100     500    1000    2000    5000   10000   25000   50000  100000999999999"
            
    oCon.AbreConexion
            
    For m = 1 To 2
       
        lsTipMon = Trim(Str(m))
        lnTCFHoy = IIf(lsTipMon = Moneda.gMonedaNacional, 3.48, 1)
        lnCntPag = 1
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        lsMoneda = IIf(m = gMonedaNacional, "SOLES", "DOLARES")
        lsTitRp1 = "E S T R A T I F I C A C I O N   D E   L O S   D E P O S I T O S"
        lsTitRp2 = ""
        lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(32, " ") & " A  H  O  R  R  O  S   P L A Z O  F I J O       C    T    S          T  O  T  A  L" & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(7, " ") & "R  A  N  G  O  S" & String(9, " ")
        lsRTFRfm = lsRTFRfm & "NUMERO   S A L D O S  NUMERO   S A L D O S  NUMERO   S A L D O S  NUMERO   S A L D O S" & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        
        For P = 1 To 10
            lnLimInd = IIf(P > 1, P - 1, P)
            If P <> 1 Then
                lnMinVal = IIf(P > 1, CCur(Mid$(lsLimite, (lnLimInd * 8) - 7, 8)), 0) * lnTCFHoy + 0.01
            Else
                lnMinVal = IIf(P > 1, CCur(Mid$(lsLimite, (lnLimInd * 8) - 7, 8)), 0) * lnTCFHoy + 0
            End If
            lnMaxVal = CCur(Mid$(lsLimite, (P * 8) - 7, 8)) * lnTCFHoy
            lsLimInf = IIf(P > 1, "DE " & JDNum(Trim(Str(lnMinVal)), 14, True, 11, 2) & " A ", "HASTA...............")
            If P > 9 Then
               lnMinVal = CCur(Mid$(lsLimite, (lnLimInd * 8) - 7, 8)) * lnTCFHoy
               lsLimInf = "MAYOR A............." & JDNum(Trim(Str(lnMinVal)), 14, True, 11, 2)
               lnMinVal = lnMinVal + 0.01
               lsLimSup = ""
            Else
               lsLimSup = JDNum(Trim(Str(lnMaxVal)), 14, True, 13, 2)
            End If
            
            lsRTFRfm = lsRTFRfm & lsLimInf & lsLimSup & "  "
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR" _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '" & Producto.gCapAhorros & "'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 14, True, 11, 2) & "  "
            lnRecs01 = lnRecs01 + rsQrySdo!nRecs:  lnSdos01 = lnSdos01 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:  lnSdoTot = lnSdoTot + lnSaldos
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR " _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '" & Producto.gCapPlazoFijo & "'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 14, True, 11, 2) & "  "
            lnRecs02 = lnRecs02 + rsQrySdo!nRecs:   lnSdos02 = lnSdos02 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:   lnSdoTot = lnSdoTot + lnSaldos
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR" _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '" & Producto.gCapCTS & "'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 14, True, 11, 2) & "  "
            lnRecs03 = lnRecs03 + rsQrySdo!nRecs:    lnSdos03 = lnSdos03 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:    lnSdoTot = lnSdoTot + lnSaldos
            
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecPro)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoTot)), 14, True, 11, 2) & Chr$(10)
            lnRecs04 = lnRecs04 + lnRecPro:    lnSdos04 = lnSdos04 + lnSdoTot
                  
            lnSdoTot = 0: lnRecPro = 0
        Next P
        
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
        lsRTFRfm = lsRTFRfm & "T  O  T  A  L  E  S" & String(17, " ")
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs01)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos01)), 14, True, 11, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs02)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos02)), 14, True, 11, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs03)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos03)), 14, True, 11, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs04)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos04)), 14, True, 11, 2) & Chr$(10)
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(10) & Chr$(12)
           
        lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0
        lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0
    Next m
    
    EstratDep = lsRTFRfm
End Function

'   ------------------------------------------------------------
'   Función     :   CabeRepo
'   Propósito   :   Define cabecera del reporte
'   Parámetro(s):   pnCarLin -> Caracteres x linea del reporte
'                   psSeccio -> Area responsable
'                   psTitRp1 -> Titulo 1 del reporte
'                   psTitRp2 -> Titulo 2 del reporte
'                   psMoneda -> Tipo de Moneda
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Formato     :   CabeRepo()
Public Function CabeRepo1(psCabe01 As String, psCabe02 As String, _
                          pnCarLin As Integer, psSeccio As String, _
                          psTitRp1 As String, psTitRp2 As String, _
                          psMoneda As Moneda, psNumPag As String) As String

    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsMoneda As String
    
    lsTitRp1 = "": lsTitRp2 = ""
    CabeRepo1 = ""
    lsMoneda = ""
    lsMoneda = IIf(psMoneda = "", String(10, " "), " - " & psMoneda)
'    psCabe01 = ""
'    psCabe02 = ""
    
    '   Definición de Cabecera 1
 
    psCabe01 = FillText(UCase(Trim(lsNomAge)) & lsMoneda, 36, " ")
    psCabe01 = psCabe01 & Space((pnCarLin - 36) - (Len(psCabe01) - 2))
    psCabe01 = psCabe01 & "PAGINA: " & psNumPag
    psCabe01 = psCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
    psCabe02 = FillText(psSeccio, 19, " ")
    psCabe02 = psCabe02 & Space((pnCarLin - 19) - (Len(psCabe02) - 2))
    psCabe02 = psCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
    lsTitRp1 = String(Int((pnCarLin - Len(psTitRp1)) / 2), " ") & psTitRp1
    lsTitRp2 = String(Int((pnCarLin - Len(psTitRp2)) / 2), " ") & psTitRp2
    
    CabeRepo1 = CabeRepo1 & psCabe01 & Chr$(10)
    CabeRepo1 = CabeRepo1 & psCabe02 & Chr$(10)
    CabeRepo1 = CabeRepo1 & lsTitRp1 & Chr$(10)
    CabeRepo1 = CabeRepo1 & lsTitRp2
        
End Function

Public Function LlenaEDAC(pnMoneda As Integer, pdFchPr1 As String, Optional pdFchPr2 As String) As String
    Dim rsQryEAC As ADODB.Recordset
    Set rsQryEAC = New ADODB.Recordset
    Dim lsQryEAC As String
    Dim ldFchIni As Date, ldFchFin As Date
    Dim lnCntMov As Integer
    Dim ldFchPro As Date
    Dim sqlAdd As String
    
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    lsRTFDat = ""
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cCodAge = '" & lsAgeCod & "'"
    End If
    
    ldFchIni = CDate(pdFchPr1)
    ldFchFin = CDate(pdFchPr2)
    
    ReDim laMovDia(1 To 14, 1 To 1)
        lnTotA02 = 0: lnTotA03 = 0: lnTotA04 = 0: lnTotA05 = 0
        lnTotA06 = 0: lnTotA07 = 0: lnTotA08 = 0: lnTotA09 = 0
        lnTotA10 = 0: lnTotA11 = 0: lnTotA12 = 0: lnTotA13 = 0
        lnTotA14 = 0
        lnCntMov = 0
        
        Do While ldFchIni <= ldFchFin
        
           lnCntMov = lnCntMov + 1
           ReDim Preserve laMovDia(1 To 14, 1 To lnCntMov)
           
           ldFchPro = ldFchIni
           
           lsQryEAC = " Select Convert(Varchar(8),dEstad,112) cUltimoMovimiento, Sum(nNumAper) nAperturasNum, Sum(nNumCanc) nCancelacionesNum," _
                    & " Sum(nMonAper) nAperturasMonto, Sum(nMonCanc) nCancelacionesMonto, Sum(nRetInt) nRetInt,  Sum(nRetInac) nRetiroIntacMon," _
                    & " Sum(nNumCancInac) nCancelacionesInacNum, Sum(nMonCancInac) nCancelacionesInacMonto, Sum(nNumDep) nDepositosNum, Sum(nMonDep) nDepositosMonto," _
                    & " Sum(nNumRet) nRetiroNum, Sum(nMonRet) nRetiroMonto, Sum(nIntCap) nInteresCapitalizados, Sum(Dia.nSaldo) nSaldo, Sum(DiaAnt.nSaldo) nSaldoAnterior," _
                    & " Sum(nMonChq) nChequeValorizacionMonto, Sum(NumCtas.NumCta) nCuentasVigentes," _
                    & " Sum(DiaSaldoCMAC.nSaldo) nSaldoCMAC, Sum(DiaSaldoCMAC.nMonChqVal) nChequeCMAC" _
                    & " From CapEstadMovimiento Dia" _
                    & " Left Join" _
                    & " (Select Convert(Varchar(8),DateAdd(Day,1,dEstad),112) Fecha, Sum(nSaldo) nSaldo From CapEstadMovimiento" _
                    & " Where nProducto = '" & Producto.gCapAhorros & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),DateAdd(Day,1,dEstad),112)) As DiaAnt On DiaAnt.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join" _
                    & " ( Select  Sum(nSaldo) nSaldo, Sum(nMonChqVal) nMonChqVal, Convert(Varchar(8),dEstad,112) Fecha"
                    
                    'lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapAhorros & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And nPersoneria = " & PersPersoneria.gPersonaJurCFLCMAC & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
                    lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapAhorros & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
                    
                    lsQryEAC = lsQryEAC & "         On DiaSaldoCMAC.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join " _
                    & " (  Select sum(nNumCtas) NumCta, cCodAge cCodAgeC, Convert(Varchar(8),dEstad,112) Fecha " _
                    & "    From CapEstadSaldo DiaSaldo where nProducto = '" & Producto.gCapAhorros & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & "  GroUp By Convert(Varchar(8),dEstad,112), cCodAge ) As NumCtas On NumCtas.Fecha = Convert(Varchar(8),Dia.dEstad,112) And Dia.cCodAge = NumCtas.cCodAgeC " _
                    & "  Where nProducto = '" & Producto.gCapAhorros & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And dEstad Between '" & Format(CDate(ldFchPro), gsFormatoFecha) & "' AND '" & Format(CDate(ldFchPro) + 1, gsFormatoFecha) & "' " & sqlAdd & " Group By  Convert(Varchar(8),dEstad,112)"

           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If rsQryEAC.EOF And rsQryEAC.BOF Then
              rsQryEAC.Close
              LlenaEDAC = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For i = 2 To 14
                  laMovDia(i, lnCntMov) = 0
              Next i
           Else
           With rsQryEAC
                laMovDia(1, lnCntMov) = Mid(!cUltimoMovimiento, 7, 2) & "/" & Mid(!cUltimoMovimiento, 5, 2) & "/" & Left(!cUltimoMovimiento, 4)
                laMovDia(2, lnCntMov) = !nAperturasNum
                laMovDia(3, lnCntMov) = !nCancelacionesNum
                laMovDia(4, lnCntMov) = !nCancelacionesInacNum
                laMovDia(5, lnCntMov) = !nCuentasVigentes
                laMovDia(6, lnCntMov) = !nDepositosNum
                laMovDia(7, lnCntMov) = !nRetiroNum
                laMovDia(8, lnCntMov) = !nAperturasMonto
                laMovDia(9, lnCntMov) = !nCancelacionesMonto
                laMovDia(10, lnCntMov) = !nCancelacionesInacMonto
                laMovDia(11, lnCntMov) = !nDepositosMonto
                laMovDia(12, lnCntMov) = !nRetiroMonto + !nRetiroIntacMon
                laMovDia(13, lnCntMov) = !nInteresCapitalizados
                laMovDia(14, lnCntMov) = (!nSaldo + !nChequeValorizacionMonto + !nChequeCMAC)
                
                lnTotA02 = lnTotA02 + !nAperturasNum
                lnTotA03 = lnTotA03 + !nCancelacionesNum
                lnTotA04 = lnTotA04 + !nCancelacionesInacNum
                lnTotA05 = lnTotA05 + !nCuentasVigentes
                lnTotA06 = lnTotA06 + !nDepositosNum
                lnTotA07 = lnTotA07 + !nRetiroNum
                lnTotA08 = lnTotA08 + !nAperturasMonto
                lnTotA09 = lnTotA09 + !nCancelacionesMonto
                lnTotA10 = lnTotA10 + !nCancelacionesInacMonto
                lnTotA11 = lnTotA11 + !nDepositosMonto
                lnTotA12 = lnTotA12 + !nRetiroMonto + !nRetiroIntacMon
                lnTotA13 = lnTotA13 + !nInteresCapitalizados
                laMovDia(14, lnCntMov) = !nSaldo
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop
        PrvEDAC pnMoneda
        
        LlenaEDAC = lsRTFDat
End Function

Public Function PrvEDAC(pnMoneda As Integer)
    Dim lnCarLin As Integer, i As Integer
    Dim lnCabece As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCntArr As Integer
    Dim lsMoneda As String
    
    Dim lsAgencia As String
    If lsAgeCod = "" Then
        lsAgencia = ""
    Else
        lsAgencia = " Agencia '" & lsAgeCod & "'"
    End If
    
    lnCabece = 0
    lnCarLin = 136
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.pAge)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        If lsAgeCod = "" Then
            lsTitRep = "MOVIMIENTO DIARIO DE AHORRO CORRIENTE CONSOLIDADO"
        Else
            lsTitRep = "MOVIMIENTO DIARIO DE AHORRO CORRIENTE - AGENCIA " & lsAgeCod
        End If
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & Chr$(10)
    lsRTFDat = lsRTFDat & lsCabe02 & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)
    lsRTFDat = lsRTFDat & "    M O V I M I E N T O S" & Space(27) & "M   O   N   T   O   S" & Chr$(10)
    lsRTFDat = lsRTFDat & "  FECHAS   "
    lsRTFDat = lsRTFDat & "APERT "
    lsRTFDat = lsRTFDat & "CANCE "
    lsRTFDat = lsRTFDat & "C.INA "
    lsRTFDat = lsRTFDat & "VIGEN "
    lsRTFDat = lsRTFDat & "DEPOS "
    lsRTFDat = lsRTFDat & "RETIR "
    lsRTFDat = lsRTFDat & "  APERTURAS "
    lsRTFDat = lsRTFDat & "   CANCELAC "
    lsRTFDat = lsRTFDat & " CANC.INACT "
    lsRTFDat = lsRTFDat & "    DEPOSITOS "
    lsRTFDat = lsRTFDat & "      RETIROS "
    lsRTFDat = lsRTFDat & "   INT.CAP. "
    lsRTFDat = lsRTFDat & "   SALDO A.C." & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)           ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For i = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, i) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, i), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(12, i), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(13, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(14, i), 13, True, 10, 2) & Chr$(10)  '  Format(laMovDia(14, I), "@@@@@@@@@@@@") & Chr$(10)
        lnTotA05 = laMovDia(5, i)
    Next i
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "  TOTAL    "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA02), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA03), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA04), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA06), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA07), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA08), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA09), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA10), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA11), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA12), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA13), 11, True, 8, 2) & Chr$(10)  '   " "
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & Chr$(10)
    lsRTFDat = lsRTFDat & "DIARIO     "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA02 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA03 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA04 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA06 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA07 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA08 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA09 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA10 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA11 / lnCntArr), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA12 / lnCntArr), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA13 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA14 / lnCntArr), 13, True, 10, 2) & Chr(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & Chr(10) & Chr(12)              ' hasta 160 caracteres x línea
End Function

Public Function LlenaEDPF(pnMoneda As Integer, pdFchPr1 As String, Optional pdFchPr2 As String) As String

    Dim rsQryEAC As New ADODB.Recordset
    Dim lsQryEAC As String
    Dim ldFchIni As Date, ldFchFin As Date
    Dim lnCntMov As Integer
    Dim ldFchPro As Date
    Dim lbCieMes As Boolean
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cCodAge = '" & lsAgeCod & "'"
    End If
    
    oCon.AbreConexion
    
    ldFchIni = CDate(pdFchPr1)
    ldFchFin = CDate(pdFchPr2)
    ReDim laMovDia(1 To 11, 1 To 1)
        lnTotA02 = 0: lnTotA03 = 0: lnTotA04 = 0: lnTotA05 = 0
        lnTotA06 = 0: lnTotA07 = 0: lnTotA08 = 0: lnTotA09 = 0
        lnTotA10 = 0: lnTotA11 = 0
        
        lnCntMov = 0
        
        Do While ldFchIni <= ldFchFin
        
           lnCntMov = lnCntMov + 1
           ReDim Preserve laMovDia(1 To 11, 1 To lnCntMov)
           
           ldFchPro = ldFchIni
           
           'sQryEAC = " Select cCodAge, Convert(Varchar(8),dEstad,112) cUltimoMovimiento, nMoneda, nNumAper nAperturasNum, nNumCanc nCancelacionesNum," _
                    & " nMonAper nAperturasMonto, nMonCanc nCancelacionesMonto, nRetInt nMovimientosMonto, nRetInac nRetiroIntacMon," _
                    & " nNumCancInac nCancelacionesInacNum, nMonCancInac, nNumDep nDepositosNum, nMonDep nDepositosMonto," _
                    & " nNumRet nRetiroNum, nMonRet nRetiroMonto, nIntCap nIntnteresCapitalizado, Dia.nSaldo, DiaAnt.nSaldo nSaldoAnterior, " _
                    & " nMonChq nChequeValorizacionMonto, cUsuario, (Select sum(nNumCtas) From CapEstadSaldo DiaSaldo where nProducto =  '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And Convert(Varchar(8),Dia.dEstad,112) = Convert(Varchar(8),DiaSaldo.dEstad,112) ) nCuentasVigentes," _
                    & " DiaSaldoCMAC.nSaldo nSaldoCMAC, DiaSaldoCMAC.nMonChqVal nChequeCMAC" _
                    & " From CapEstadMovimiento Dia" _
                    & " Inner Join" _
                    & " (Select Convert(Varchar(8),DateAdd(Day,1,dEstad),112) Fecha, nSaldo From CapEstadMovimiento" _
                    & " Where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda =  " & Trim(Str(pnMoneda)) & "  " & sqlAdd & ") As DiaAnt On DiaAnt.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join" _
                    & " ( Select  nSaldo , nMonChqVal, Convert(Varchar(8),DateAdd(Day,1,dEstad),112) Fecha" _
                    & "     From CapEstadSaldo where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda =  " & Trim(Str(pnMoneda)) & " And nPersoneria = " & PersPersoneria.gPersonaJurCFLCMAC & " " & sqlAdd & ") As DiaSaldoCMAC" _
                    & "      On DiaSaldoCMAC.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & "  Where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And dEstad Between '" & Format(CDate(ldFchPro), gsFormatoFecha) & "' AND '" & Format(CDate(ldFchPro) + 1, gsFormatoFecha) & "' " & sqlAdd
           
           lsQryEAC = " Select Convert(Varchar(8),dEstad,112) cUltimoMovimiento, Sum(nNumAper) nAperturasNum, Sum(nNumCanc) nCancelacionesNum," _
                    & " Sum(nMonAper) nAperturasMonto, Sum(nMonCanc) nCancelacionesMonto, Sum(nRetInt) nMovimientosMonto,  Sum(nRetInac) nRetiroIntacMon," _
                    & " Sum(nNumCancInac) nCancelacionesInacNum, Sum(nMonCancInac) nCancelacionesInacMonto, Sum(nNumDep) nDepositosNum, Sum(nMonDep) nDepositosMonto," _
                    & " Sum(nNumRet) nRetiroNum, Sum(nMonRet) nRetiroMonto, Sum(nIntCap) nInteresCapitalizados, Sum(Dia.nSaldo) nSaldo, Sum(DiaAnt.nSaldo) nSaldoAnterior," _
                    & " Sum(nMonChq) nChequeValorizacionMonto, Sum(NumCtas.NumCta) nCuentasVigentes," _
                    & " Sum(DiaSaldoCMAC.nSaldo) nSaldoCMAC, Sum(DiaSaldoCMAC.nMonChqVal) nChequeCMAC" _
                    & " From CapEstadMovimiento Dia" _
                    & " Left Join" _
                    & " (Select Convert(Varchar(8),DateAdd(Day,1,dEstad),112) Fecha, Sum(nSaldo) nSaldo From CapEstadMovimiento" _
                    & " Where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),DateAdd(Day,1,dEstad),112)) As DiaAnt On DiaAnt.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join" _
                    & " ( Select  Sum(nSaldo) nSaldo, Sum(nMonChqVal) nMonChqVal, Convert(Varchar(8),dEstad,112) Fecha"
                    
            'lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And nPersoneria = " & PersPersoneria.gPersonaJurCFLCMAC & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
            lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
            
            lsQryEAC = lsQryEAC & "         On DiaSaldoCMAC.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join " _
                    & " (  Select sum(nNumCtas) NumCta, cCodAge cCodAgeC, Convert(Varchar(8),dEstad,112) Fecha " _
                    & "    From CapEstadSaldo DiaSaldo where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & "  GroUp By Convert(Varchar(8),dEstad,112), cCodAge ) As NumCtas On NumCtas.Fecha = Convert(Varchar(8),Dia.dEstad,112) And Dia.cCodAge = NumCtas.cCodAgeC " _
                    & "  Where nProducto = '" & Producto.gCapPlazoFijo & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And dEstad Between '" & Format(CDate(ldFchPro), gsFormatoFecha) & "' AND '" & Format(CDate(ldFchPro) + 1, gsFormatoFecha) & "' " & sqlAdd & " Group By  Convert(Varchar(8),dEstad,112)"
           
           
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If RSVacio(rsQryEAC) Then
              rsQryEAC.Close
              Set rsQryEAC = Nothing
              LlenaEDPF = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For i = 2 To 11
                  laMovDia(i, lnCntMov) = 0
              Next i
           Else
           
           With rsQryEAC
                           
                laMovDia(1, lnCntMov) = Mid(!cUltimoMovimiento, 7, 2) & "/" & Mid(!cUltimoMovimiento, 5, 2) & "/" & Left(!cUltimoMovimiento, 4)
                laMovDia(2, lnCntMov) = !nAperturasNum
                laMovDia(3, lnCntMov) = !nCancelacionesNum
                laMovDia(4, lnCntMov) = !nRetiroNum
                laMovDia(5, lnCntMov) = !nCuentasVigentes
                laMovDia(6, lnCntMov) = !nAperturasMonto
                laMovDia(7, lnCntMov) = !nCancelacionesMonto
                laMovDia(8, lnCntMov) = !nMovimientosMonto
                laMovDia(9, lnCntMov) = !nInteresCapitalizados
                laMovDia(10, lnCntMov) = !nSaldo
                laMovDia(11, lnCntMov) = 0
                
                lnTotA02 = lnTotA02 + !nAperturasNum
                lnTotA03 = lnTotA03 + !nCancelacionesNum
                lnTotA04 = lnTotA04 + !nRetiroNum
                lnTotA05 = lnTotA05 + !nCuentasVigentes
                lnTotA06 = lnTotA06 + !nAperturasMonto
                lnTotA07 = lnTotA07 + !nCancelacionesMonto
                lnTotA08 = lnTotA08 + !nMovimientosMonto
                lnTotA09 = lnTotA09 + !nInteresCapitalizados
                lnTotA10 = lnTotA10 + !nSaldo
                lnTotA11 = 0
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop
        
        If CierreRealizado(2) Then
           lbCieMes = True
        Else
           lbCieMes = False
        End If
        
        If lbCieMes Then
           lsQryEAC = "SELECT SUM(ABS(nIntAcum)) AS nIntDev FROM Captaciones CA Inner Join Producto PR ON CA.cCtaCod = PR.cCtaCod " _
                    & "WHERE SUBSTRING(CA.cCtaCod,9,1) = '" & Trim(Str(pnMoneda)) & "' AND Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Substring(CA.cCtaCod,6,3) = '233' "
           
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
        
           If RSVacio(rsQryEAC) Then
              laMovDia(11, lnCntMov) = 0
              lnTotA11 = 0
           Else
              laMovDia(11, lnCntMov) = IIf(IsNull(rsQryEAC!nIntDev), 0, rsQryEAC!nIntDev)
              lnTotA11 = IIf(IsNull(rsQryEAC!nIntDev), 0, rsQryEAC!nIntDev)
           End If
        End If
        
        PrvEDPF pnMoneda
        LlenaEDPF = lsRTFDat
End Function

Public Function PrvEDPF(pnMoneda As Integer)
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCabece As Integer
    Dim lnCntArr As Integer
    
    Dim lsAgencia As String
    If lsAgeCod = "" Then
        lsAgencia = ""
    Else
        lsAgencia = " Agencia '" & lsAgeCod & "'"
    End If
    
    lnCabece = 0
    lnCarLin = 110
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.pAge)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        If lsAgeCod = "" Then
            lsTitRep = "MOVIMIENTO DIARIO DE PLAZO FIJO CONSOLIDADO"
        Else
            lsTitRep = "MOVIMIENTO DIARIO DE PLAZO FIJO - AGENCIA " & lsAgeCod
        End If
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & Chr$(10)
    lsRTFDat = lsRTFDat & lsCabe02 & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)
    lsRTFDat = lsRTFDat & "       M O V I M I E N T O S" & Space(21) & "M   O   N   T   O   S" & String(16, " ")
    lsRTFDat = lsRTFDat & "  SALDOS PF " & Chr$(10)
    lsRTFDat = lsRTFDat & "  FECHAS   "
    lsRTFDat = lsRTFDat & "APERT "
    lsRTFDat = lsRTFDat & "CANCE "
    lsRTFDat = lsRTFDat & "MOVIM "
    lsRTFDat = lsRTFDat & "VIGEN "
    lsRTFDat = lsRTFDat & "  APERTURAS "
    lsRTFDat = lsRTFDat & "   CANCELAC "
    lsRTFDat = lsRTFDat & " MOVIMIENTO "
    lsRTFDat = lsRTFDat & "   INT.CAP. "
    lsRTFDat = lsRTFDat & "      CAPITAL "
    lsRTFDat = lsRTFDat & "  INT.DEVENG." & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)           ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For i = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, i) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, i), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, i), 13, True, 10, 2) & Chr$(10)
        lnTotA05 = laMovDia(5, i)
    Next i
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "  TOTAL    "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA02), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA03), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA04), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA06), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA07), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA08), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA09), 11, True, 8, 2) & Chr$(10)
'    lsRTFDat = lsRTFDat & JDNum(lnTotA10, 12, True) & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & Chr$(10)
    lsRTFDat = lsRTFDat & "DIARIO     "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA02 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA03 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA04 / lnCntArr), 5, False, 5, 0) & " "
'    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05 / lnCntArr), 4, False, 4, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA06 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA07 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA08 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA09 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA10 / lnCntArr), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA11 / lnCntArr), 13, True, 10, 2) & Chr(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & Chr$(10) & Chr(12)              ' hasta 160 caracteres x línea
End Function

Public Function LlenaEDCTS(pnMoneda As Integer, pdFchPr1 As String, Optional pdFchPr2 As String) As String
    
    Dim rsQryEAC As New ADODB.Recordset
    Dim lsQryEAC As String
    Dim ldFchIni As Date, ldFchFin As Date
    Dim lnCntMov As Integer
    Dim ldFchPro As Date
    Dim lbCieMes As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    ldFchIni = CDate(pdFchPr1)
    ldFchFin = CDate(pdFchPr2)
    
    oCon.AbreConexion
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cCodAge = '" & lsAgeCod & "'"
    End If
    
    ReDim laMovDia(1 To 14, 1 To 1)
    
        lnTotA02 = 0: lnTotA03 = 0: lnTotA04 = 0: lnTotA05 = 0
        lnTotA06 = 0: lnTotA07 = 0: lnTotA08 = 0: lnTotA09 = 0
        lnTotA10 = 0: lnTotA11 = 0: lnTotA12 = 0: lnTotA13 = 0
        lnTotA14 = 0
        lnCntMov = 0
        
        Do While ldFchIni <= ldFchFin
        
           lnCntMov = lnCntMov + 1
           ReDim Preserve laMovDia(1 To 14, 1 To lnCntMov)
           
           ldFchPro = ldFchIni
           
           lsQryEAC = " SELECT cUltimoMovimiento,nMoneda,nAperturasNum,nAperturasMonto,nCancelacionesNum,nCancelacionesMonto,nRetiroNum,nRetiroMonto,nRetiroIntCTS,nDepositosNum,nDepositosMonto,nCuentasVigentes,nInteresCapitalizado,nSaldoCTS,nSaldEmp,nChequeMonto,nChequeValorizacionMonto,nChequevalorizacionNum FROM Estadisticacts" _
                    & " WHERE nMoneda = '" & Trim(Str(pnMoneda)) & "' AND cUltimoMovimiento BETWEEN '" & Format(CDate(ldFchPro), "yyyymmdd") & "' AND '" & Format(CDate(ldFchPro) + 1, "yyyymmdd") & "' AND Substring(cUltimoMovimiento,18,2) = '" & lsAgeCod & "' Order by cUltimoMovimiento "
           
           
           lsQryEAC = " Select Convert(Varchar(8),dEstad,112) cUltimoMovimiento, Sum(nNumAper) nAperturasNum, Sum(nNumCanc) nCancelacionesNum," _
                    & " Sum(nMonAper) nAperturasMonto, Sum(nMonCanc) nCancelacionesMonto, Sum(nRetInt) nRetInt,  Sum(nRetInac) nRetiroIntacMon," _
                    & " Sum(nNumCancInac) nCancelacionesInacNum, Sum(nMonCancInac) nCancelacionesInacMonto, Sum(nNumDep) nDepositosNum, Sum(nMonDep) nDepositosMonto," _
                    & " Sum(nNumRet) nRetiroNum, Sum(nMonRet) nRetiroMonto, Sum(nIntCap) nInteresCapitalizados, Sum(Dia.nSaldo) nSaldo, Sum(DiaAnt.nSaldo) nSaldoAnterior," _
                    & " Sum(nMonChq) nChequeValorizacionMonto, Sum(NumCtas.NumCta) nCuentasVigentes," _
                    & " Sum(DiaSaldoCMAC.nSaldo) nSaldoCMAC, Sum(DiaSaldoCMAC.nMonChqVal) nChequeCMAC, Sum(NumCtas.SaldoEmp) SaldoEmp" _
                    & " From CapEstadMovimiento Dia" _
                    & " Left Join" _
                    & " (Select Convert(Varchar(8),DateAdd(Day,1,dEstad),112) Fecha, Sum(nSaldo) nSaldo From CapEstadMovimiento" _
                    & " Where nProducto = '" & Producto.gCapCTS & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),DateAdd(Day,1,dEstad),112)) As DiaAnt On DiaAnt.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Left Join" _
                    & " ( Select  Sum(nSaldo) nSaldo, Sum(nMonChqVal) nMonChqVal, Convert(Varchar(8),dEstad,112) Fecha"
                    
            'lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapCTS & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And nPersoneria = " & PersPersoneria.gPersonaJurCFLCMAC & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
            lsQryEAC = lsQryEAC & "     From CapEstadSaldo where nProducto = '" & Producto.gCapCTS & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & " Group By Convert(Varchar(8),dEstad,112)) As DiaSaldoCMAC"
            
            lsQryEAC = lsQryEAC & "         On DiaSaldoCMAC.Fecha = Convert(Varchar(8),dEstad,112)" _
                    & " Inner Join " _
                    & " ( Select sum(nNumCtas) NumCta, sum(nSaldoInactEmp) SaldoEmp, cCodAge cCodAgeC, Convert(Varchar(8),dEstad,112) Fecha " _
                    & "   From CapEstadSaldo DiaSaldo where nProducto = '" & Producto.gCapCTS & "' And nMoneda = " & Trim(Str(pnMoneda)) & " " & sqlAdd & "  GroUp By Convert(Varchar(8),dEstad,112), cCodAge ) As NumCtas On NumCtas.Fecha = Convert(Varchar(8),Dia.dEstad,112) And Dia.cCodAge = NumCtas.cCodAgeC " _
                    & "   Where nProducto = '" & Producto.gCapCTS & "' And nMoneda = " & Trim(Str(pnMoneda)) & " And dEstad Between '" & Format(CDate(ldFchPro), gsFormatoFecha) & "' AND '" & Format(CDate(ldFchPro) + 1, gsFormatoFecha) & "' " & sqlAdd & " Group By  Convert(Varchar(8),dEstad,112)"
           
           
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If RSVacio(rsQryEAC) Then
              rsQryEAC.Close
              Set rsQryEAC = Nothing
              LlenaEDCTS = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For i = 2 To 14
                  laMovDia(i, lnCntMov) = 0
              Next i
'              Exit Function
           Else
           
           With rsQryEAC
                           
                laMovDia(1, lnCntMov) = Mid(!cUltimoMovimiento, 7, 2) & "/" & Mid(!cUltimoMovimiento, 5, 2) & "/" & Left(!cUltimoMovimiento, 4)
                laMovDia(2, lnCntMov) = !nAperturasNum
                laMovDia(3, lnCntMov) = !nCancelacionesNum
                laMovDia(4, lnCntMov) = !nCuentasVigentes
                laMovDia(5, lnCntMov) = !nDepositosNum
                laMovDia(6, lnCntMov) = !nRetiroNum
                laMovDia(7, lnCntMov) = !nAperturasMonto
                laMovDia(8, lnCntMov) = !nCancelacionesMonto
                laMovDia(9, lnCntMov) = !nDepositosMonto
                laMovDia(10, lnCntMov) = !nRetiroMonto
                laMovDia(11, lnCntMov) = !nInteresCapitalizados
                laMovDia(12, lnCntMov) = !nSaldo
                laMovDia(13, lnCntMov) = !SaldoEmp
                laMovDia(14, lnCntMov) = 0
                
                lnTotA02 = lnTotA02 + !nAperturasNum
                lnTotA03 = lnTotA03 + !nCancelacionesNum
                lnTotA04 = lnTotA04 + !nCuentasVigentes
                lnTotA05 = lnTotA05 + !nDepositosNum
                lnTotA06 = lnTotA06 + !nRetiroNum
                lnTotA07 = lnTotA07 + !nAperturasMonto
                lnTotA08 = lnTotA08 + !nCancelacionesMonto
                lnTotA09 = lnTotA09 + !nDepositosMonto
                lnTotA10 = lnTotA10 + !nRetiroMonto
                lnTotA11 = lnTotA11 + !nInteresCapitalizados
                lnTotA12 = lnTotA12 + !nSaldo
                lnTotA13 = lnTotA13 + !SaldoEmp
                lnTotA14 = 0
                
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop

        If CierreRealizado(2) Then
           lbCieMes = True
        Else
           lbCieMes = False
        End If
        
        If lbCieMes Then
            ldFchIni = CDate(pdFchPr1)
            ldFchFin = CDate(pdFchPr2)
            lsQryEAC = " SELECT SUM(ABS(nMonTran)) AS nMonPro FROM TransAho " _
                     & " WHERE  cCodOpe = '" & 333 & "' AND (cFlag IS NULL OR cFlag <> 'X') " _
                     & " and    substring(cCodCta,6,1) = '" & Trim(Str(pnMoneda)) & "' " _
                     & " AND    (dFecTran BETWEEN '" & Format(ldFchIni, gsFormatoFecha) & "' AND '" & Format(DateAdd("d", 1, ldFchFin), gsFormatoFecha) & "') "
           
           'rsQryEAC.Open lsQryEAC, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
           
           If RSVacio(rsQryEAC) Then
              laMovDia(14, lnCntMov) = 0
              lnTotA14 = 0
           Else
              laMovDia(14, lnCntMov) = IIf(IsNull(rsQryEAC!nMonPro), 0, rsQryEAC!nMonPro)
              lnTotA14 = IIf(IsNull(rsQryEAC!nMonPro), 0, rsQryEAC!nMonPro)
           End If
        End If
        
        PrvEDCTS pnMoneda
        LlenaEDCTS = lsRTFDat

       
End Function

Public Function PrvEDCTS(pnMoneda As Integer)
    
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCabece As Integer
    Dim lnCntArr As Integer
    
    Dim lsAgencia As String
    If lsAgeCod = "" Then
        lsAgencia = ""
    Else
        lsAgencia = " Agencia '" & lsAgeCod & "'"
    End If
    
    lnCabece = 0
    lnCarLin = 142
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.pAge)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        If lsAgeCod = "" Then
            lsTitRep = "MOVIMIENTO DIARIO DE C T S - CONSOLIDADO"
        Else
            lsTitRep = "MOVIMIENTO DIARIO DE C T S - AGENCIA " & lsAgeCod
        End If
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & Chr$(10)
    lsRTFDat = lsRTFDat & lsCabe02 & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & Chr$(10) & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)
    lsRTFDat = lsRTFDat & "    M O V I M I E N T O S" & Space(27) & "M   O   N   T   O   S" & Chr$(10)
    lsRTFDat = lsRTFDat & "  FECHAS   "
    lsRTFDat = lsRTFDat & "APERT "
    lsRTFDat = lsRTFDat & "CANCE "
    lsRTFDat = lsRTFDat & "VIGEN "
    lsRTFDat = lsRTFDat & "DEPOS "
    lsRTFDat = lsRTFDat & "RETIR "
    lsRTFDat = lsRTFDat & "  APERTURAS "
    lsRTFDat = lsRTFDat & "   CANCELAC "
    lsRTFDat = lsRTFDat & "  DEPOSITOS "
    lsRTFDat = lsRTFDat & "    RETIROS "
    lsRTFDat = lsRTFDat & "   INT.CAP. "
    lsRTFDat = lsRTFDat & "    SALDO CTS"
    lsRTFDat = lsRTFDat & "  SALDO EMPLEA"
    lsRTFDat = lsRTFDat & "   PROVISIONES" & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)           ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For i = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, i) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, i), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, i), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(12, i), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(13, i), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(14, i), 13, True, 10, 2) & Chr$(10)  '  Format(laMovDia(14, I), "@@@@@@@@@@@@") & Chr$(10)
        lnTotA05 = laMovDia(5, i)
    Next i
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "  TOTAL    "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA02), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA03), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA04), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA06), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA07), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA08), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA09), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA10), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA11), 11, True, 8, 2) & Chr$(10)
'    lsRTFDat = lsRTFDat & Format(Format(lnTotA13, "###,##0.00"), "@@@@@@@@@@") & Chr$(10)  '   " "
'    lsRTFDat = lsRTFDat & Format(Format(lnTotA14, "#,###,##0.00"), "@@@@@@@@@@@@") & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & Chr$(10)
    lsRTFDat = lsRTFDat & "DIARIO     "
    lsRTFDat = lsRTFDat & JDNum((lnTotA02 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA03 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA04 / lnCntArr), 5, False, 5, 0) & " "
'    lsRTFDat = lsRTFDat & JDNum((lnTotA05 / lnCntArr),4, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA06 / lnCntArr), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA07 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA08 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA09 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA10 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA11 / lnCntArr), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA12 / lnCntArr), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA13 / lnCntArr), 13, True, 10, 2) & " "
    lsRTFDat = lsRTFDat & JDNum((lnTotA14 / lnCntArr), 13, True, 10, 2) & Chr$(10)
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & Chr$(10) & Chr(12)             ' hasta 160 caracteres x línea
End Function


Public Function GenSdosFM(psTipAho As Producto, psTipPer As PersPersoneria, psTipMon As Moneda, _
                             psFchPro As String, Optional pbCtaIna As Byte = 0, _
                             Optional psCodIns As String = "") As String
    Dim lsQrySdo As String
    Dim rsQrySdo As ADODB.Recordset
    Set rsQrySdo = New ADODB.Recordset
    Dim lsInstit As String
    Dim lsMoneda As String
    Dim lsPerson As String
    Dim lsTipAho As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And PR.cCtaCod Like '___" & lsAgeCod & "%'"
    End If
    
    oCon.AbreConexion
    
    lsRTFRfm = ""
    
    lsMoneda = IIf(psTipMon = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    
    Select Case psTipPer
        Case PersPersoneria.gPersonaNat
             lsPerson = "PERSONA(S)  NATURAL(ES)" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurSFL
             lsPerson = "PERSONA(S)  JURIDICA(S)  SIN  FINES  DE  LUCRO" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurCFL
             lsPerson = "PERSONA(S)  JURIDICA(S)  CON  FINES  DE  LUCRO" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurCFLCMAC
             lsPerson = "PERSONA(S)  CMAC" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurCFLCRAC
             lsPerson = "PERSONA(S)  CRAC" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurCFLCooperativa
             lsPerson = "PERSONA(S)  COOPERATIVA(S)" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case PersPersoneria.gPersonaJurCFLEdpyme
             lsPerson = "PERSONA(S)  EDPYME(S)" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
    End Select
        
    lsQrySdo = ""
    Select Case psTipAho
        Case Producto.gCapAhorros
   
             lsTipAho = "A H O R R O S"
             lsQrySdo = " SELECT PR.cCtaCod cCodCta, bOrdPag cOrdPag, CA.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt, dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') As nIntMes, CAA.nSaldoAnterior AS nSdoAnt, CAA.dUltContacto AS dUltMovAC, PE.cPersnombre cNomPers, PP.cPersCod cCodPers FROM Producto PR" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod" _
                      & " INNER JOIN Persona PE ON Pp.cPersCod = PE.cPersCod" _
                      & " INNER JOIN Captaciones CA ON PR.cCtaCod = CA.cCtaCod" _
                      & " INNER JOIN CaptacAhorros CAA ON PR.cCtaCod = CAA.cCtaCod" _
                      & "" _
                      & " WHERE Left(PR.nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND PE.nPersPersoneria = '" & psTipPer & "'" _
                      & " AND PP.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & " AND CAA.bInactiva = " & pbCtaIna & "" _
                      & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & "" _
                      & " ORDER BY AC.cOrdPag, AC.cCodCta"
        Case Producto.gCapPlazoFijo
             lsTipAho = "P L A Z O   F I J O"
             lsQrySdo = " SELECT CA.dApertura dAperPF, DateAdd(day, CAPF.nPlazo, CAPF.dRenovacion) AS dFchVct, CAPF.nPlazo, PR.cCtaCod cCodCta, PE.cPersNombre cNomPers, CA.nSaldoDisp AS nCapital, " _
                      & " CAPF.nIntPag, CA.nIntAcum, CAPF.nApertura, PR.nTasaInteres nTasaIntPF, CAPF.dAuxiliar AS dUltMovPF, PE.cPersCod " _
                      & " FROM Producto PR" _
                      & " Inner Join CaptacPlazoFijo CAPF On PR.cCtaCod = CAPF.cCtaCod INNER JOIN Captaciones CA ON CAPF.cCtaCod = CA.cCtaCod " _
                      & " INNER JOIN ProductoPersona PP ON CAPF.cCtaCod = PP.cCtaCod INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " _
                      & " WHERE Left(PR.nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND PE.nPersPersoneria = '" & psTipPer & "' " _
                      & " AND PP.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & " AND SUBSTRING(CAPF.cCtaCod,9,1) =  '" & psTipMon & "'" _
                      & " AND SUBSTRING(CAPF.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & " " _
                      & " ORDER BY PR.cCtaCod"
        Case Producto.gCapCTS
             lsTipAho = "C T S"
             Select Case psCodIns
                Case gsInstCmac
                     lsPerson = "E M P L E A D O S    C M A C T"
                     lsInstit = "LTRIM(RTRIM(CTS.cCodInst)) = '" & gsInstCmac & "' "
                Case Else
                     lsPerson = "C L I E N T E S    C M A C T"
                     lsInstit = "LTRIM(RTRIM(CTS.cCodInst)) <> '" & gsInstCmac & "' "
             End Select
             lsQrySdo = " SELECT CA.dApertura dAperCTS, PR.cCtaCod cCodCta, CACTS.nSaldRetiro As nSdoDis, PR.nSaldo AS nSdoCnt, " _
                      & " CA.dUltCierre AS dUltMovCTS, 0 As nUltDep, PR.nSaldo As nCapital, dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers " _
                      & " FROM Producto PR  " _
                      & " INNER JOIN Captaciones CA ON CA.cCtaCod = PR.cCtaCod  INNER JOIN CaptacCTS CACTS ON CACTS.cCtaCod = PR.cCtaCod" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod" _
                      & " WHERE Left(PR.nPRDEstado,2) NOT IN ('13','14') AND LTRIM(RTRIM(CACTS.cCodInst)) = '" & lsInstit & "'" _
                      & " AND PP.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & " AND SUBSTRING(CACTS.cCtaCod,9,1) = '" & psTipPer & "'" _
                      & " AND SUBSTRING(CACTS.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & " " _
                      & " ORDER BY PR.cCtaCod"
    
    End Select
    
    If lsQrySdo = "" Then
       Exit Function
    End If
    
    Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
    
    If rsQrySdo.EOF And rsQrySdo.BOF Then
       lsRTFRfm = ""
       lsReporte = ""
       GenSdosFM = lsReporte
    Else
       lsRTFRfm = ""
       lsReporte = ""
       PrtSdosFM rsQrySdo, psTipMon, lsTipAho, lsPerson
       GenSdosFM = lsReporte
    End If
    rsQrySdo.Close
    Set rsQrySdo = Nothing
End Function


Private Sub PrtSdosFM(ByVal prQryDat As ADODB.Recordset, psMoneda As Moneda, psTipAho As String, psPerson As String)

    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Currency, lnTotRet As Currency
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Currency, lnSdoDis As Currency
    Dim lnSumInt As Currency, lnTotInt As Currency
    Dim lnCapita As Currency, lsCodCta As String
    Dim lnCapiFi As Currency, lnSdDiFi As Currency, lnSdCnFi As Currency
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnCntRel As Integer, lbCtaIde As Boolean
    Dim lnCalInt As Currency, lnNumDia As Integer
    Dim lnApertu As Currency       '   , lnCapita As Currency
    Dim lnIntDev As Currency, lnIntCal As Currency
    Dim lnTotalP As Currency, lnTotApe As Currency
    Dim lnTotCap As Currency, lnToInDv As Currency
    Dim lnToInCa As Currency, lnTotalF As Currency
    Dim lbCieMes As Boolean, lnIntCap As Currency
    
    If CierreRealizado(2) Then
       lbCieMes = True
    Else
       lbCieMes = False
    End If
    
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 112
    lsTitRp1 = "L I S T A D O   D E   C U E N T A S   D E   " & psTipAho
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
    lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
    lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
    
    lnCntPag = 0
    lbCtaIde = False
    With prQryDat
         Do While Not .EOF
            Select Case Mid$(!cCodCta, 6, 3)
               Case Producto.gCapAhorros
                    lsTitRp2 = IIf(Not !cOrdPag, "", " CON ORDEN DE PAGO  -  ") & psPerson
               Case Producto.gCapPlazoFijo
                    lsTitRp2 = psPerson
                    lnCarLin = 167
               Case Producto.gCapCTS
                    lsTitRp2 = psPerson
                    lnCarLin = 173
            End Select
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  If Mid$(!cCodCta, 6, 3) = Producto.gCapAhorros Then
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                     lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
                  Else
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & Chr$(10)
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & Chr$(10)
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
                     lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
                  End If
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     If Mid$(!cCodCta, 6, 3) = Producto.gCapAhorros Then
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                        lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                     Else
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & Chr$(10)
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                        lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
                     End If
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, IIf(psMoneda = gMonedaExtranjera, "DOLARES", "SOLES"), lsNumPag) & Chr$(10)
               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)
               
               If Mid$(!cCodCta, 6, 3) = Producto.gCapAhorros Then
                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "  CUENTA  " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "     INTERES  " & " "
                  lsRTFRfm = lsRTFRfm & "  SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & " ULTIMO  " & Chr$(10)
                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "      DEL CLIENTE     " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "    DEL MES  " & " "
                  lsRTFRfm = lsRTFRfm & "DISPONIBLE " & " "
                  lsRTFRfm = lsRTFRfm & "   CONTABLE  " & " "
                  lsRTFRfm = lsRTFRfm & "MOVIMIENTO" & Chr$(10)
               End If
               If Mid$(!cCodCta, 6, 3) = Producto.gCapPlazoFijo Then
'                  lnCarLin = 153
                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
                  lsRTFRfm = lsRTFRfm & "PLAZOS" & " "
                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & " CUENTA  " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "      DEPOSITO  " & " "
                  lsRTFRfm = lsRTFRfm & "            " & " "
                  lsRTFRfm = lsRTFRfm & "  INTERES   " & " "
                  lsRTFRfm = lsRTFRfm & "  INTERES   " & " "
                  lsRTFRfm = lsRTFRfm & "            " & " "
                  lsRTFRfm = lsRTFRfm & " ULTIMO" & Chr$(10)
                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
                  lsRTFRfm = lsRTFRfm & " APERTURA " & " "
                  lsRTFRfm = lsRTFRfm & "  VENCIM  " & " "
                  lsRTFRfm = lsRTFRfm & "(DIAS)" & " "
                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "     DEL CLIENTE     " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "    APERTURA  " & " "
                  lsRTFRfm = lsRTFRfm & "  CAPITAL   " & " "
                  lsRTFRfm = lsRTFRfm & " DEVENGADO  " & " "
                  lsRTFRfm = lsRTFRfm & " CALCULADO  " & " "
                  lsRTFRfm = lsRTFRfm & "  T O T A L " & " "
                  lsRTFRfm = lsRTFRfm & "CONTACTO" & Chr$(10)
               End If
               If Mid$(!cCodCta, 6, 3) = Producto.gCapCTS Then
                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
                  lsRTFRfm = lsRTFRfm & "PLAZOS" & " "
                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & " CUENTA  " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "   ULTIMO   " & " "
                  lsRTFRfm = lsRTFRfm & "            " & " "
                  lsRTFRfm = lsRTFRfm & "   INTERES  " & " "
                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & "    ULTIMO" & Chr$(10)
                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
                  lsRTFRfm = lsRTFRfm & " APERTURA " & " "
                  lsRTFRfm = lsRTFRfm & "  VENCIM  " & " "
                  lsRTFRfm = lsRTFRfm & "(DIAS)" & " "
                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "     DEL CLIENTE     " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "  DEPOSITO  " & " "
                  lsRTFRfm = lsRTFRfm & "  CAPITAL   " & " "
                  lsRTFRfm = lsRTFRfm & "  DEL MES   " & " "
                  lsRTFRfm = lsRTFRfm & "  DISPONIBLE" & " "
                  lsRTFRfm = lsRTFRfm & "  CONTABLE  " & " "
                  lsRTFRfm = lsRTFRfm & "  CONTACTO" & Chr$(10)
               End If
               
               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10) & Chr$(10)         ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            If !cCodCta = lsCodCta Then
                If lbCtaIde = False Then
                    Select Case Mid$(lsCodCta, 6, 3)
                        Case Producto.gCapAhorros
                            lsRTFRfm = lsRTFRfm & String(26, " ")
                            lnCntRel = lnCntRel + 1
                        Case Producto.gCapPlazoFijo
                            lsRTFRfm = lsRTFRfm & String(55, " ")
                            lnCntRel = lnCntRel + 1
                        Case Producto.gCapCTS
                            lsRTFRfm = lsRTFRfm & String(55, " ")
                            lnCntRel = lnCntRel + 1
                    End Select
                    lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & Chr(10)
                    lbCtaIde = True
               End If
            Else
               lsCodCta = !cCodCta
               lbCtaIde = False
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               Select Case Mid$(!cCodCta, 6, 3)
                  Case Producto.gCapAhorros
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       lsRTFRfm = lsRTFRfm & String(9, " ")
                       If Len(Trim(!cNomPers)) < 34 Then
                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 34), 34, " ") & " "
                       Else
                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 34) & " "
                       End If
                       lnIntCap = 0
                       If lbCieMes Then
                          lnIntCap = !nSdoDis - !nSdoAnt
                       Else
                          lnIntCap = !nIntMes
                       End If
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoCnt, 12, True, 9, 2) & "   "
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovAC, gsFormatoFechaView) & Chr(10)
                       
                       lnCapita = lnCapita + lnIntCap: lnCapiFi = lnCapiFi + lnIntCap
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapPlazoFijo
                       lnCalInt = 0
                       lnNumDia = 0
                       lnNumDia = DateDiff("d", !dUltMovPF, !dFchVct) - 1
                       lnCalInt = GetInteres(!nTasaIntPF, (!nIntAcum + !nCapital), lnNumDia, False)
                       
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
                       lsRTFRfm = lsRTFRfm & Format(!dAperPF, gsFormatoFechaView) & " "
                       lsRTFRfm = lsRTFRfm & Format(!dFchVct, gsFormatoFechaView) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nPlazo)), 6, False, 6, 0) & " "
                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       lsRTFRfm = lsRTFRfm & String(9, " ")
                       If Len(Trim(!cNomPers)) < 34 Then
                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 34), 34, " ") & " "
                       Else
                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 34) & " "
                       End If
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nApertura)), 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nCapital)), 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nIntAcum)), 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCalInt)), 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nCapital + lnCalInt)), 12, True, 9, 2) & "   "
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovPF, gsFormatoFechaView) & Chr$(10)
                       
                       lnApertu = lnApertu + !nApertura: lnTotApe = lnTotApe + !nApertura
                       lnCapita = lnCapita + !nCapital: lnTotCap = lnTotCap + !nCapital
                       lnIntDev = lnIntDev + !nIntAcum: lnToInDv = lnToInDv + !nIntAcum
                       lnIntCal = lnIntCal + lnCalInt: lnToInCa = lnToInCa + lnCalInt
                       lnTotalP = lnTotalP + (!nCapital + lnCalInt)
                       lnTotalF = lnTotalF + (!nCapital + lnCalInt)
                       
                       lnCntRel = RelCta(!cCodCta, !cPersCod)
                       lbCtaIde = True
                  Case Producto.gCapCTS
                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
                       lsRTFRfm = lsRTFRfm & Format(!dAperCTS, gsFormatoFechaView) & " "
                       lsRTFRfm = lsRTFRfm & String(10, " ") & " "
                       lsRTFRfm = lsRTFRfm & String(6, " ") & " "
                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       lsRTFRfm = lsRTFRfm & String(9, " ")
                       If Len(Trim(PstaNombre(!cNomPers, False))) < 40 Then
                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40), 40, " ") & " "
                       Else
                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & " "
                       End If
                       lsRTFRfm = lsRTFRfm & JDNum(!nUltDep, 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nCapital, 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nIntMes, 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoCnt, 12, True, 9, 2) & "   "
                     '  lsRTFRfm = lsRTFRfm & JDNum(!CHEQUESENV, 12, True, 9, 2) & "   "
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovCTS, gsFormatoFechaView) & Chr(10)
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       
                       lnApertu = lnApertu + !nUltDep: lnTotApe = lnTotApe + !nUltDep
                       lnCapita = lnCapita + !nCapital: lnTotCap = lnTotCap + !nCapital
                       lnIntDev = lnIntDev + !nIntMes: lnToInDv = lnToInDv + !nIntMes
                       lnIntCal = lnIntCal + !nSdoDis: lnToInCa = lnToInCa + !nSdoDis
                       lnTotalP = lnTotalP + !nSdoCnt: lnTotalF = lnTotalF + !nSdoCnt
                       
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
               End Select
            End If
            lnCntCta = lnCntCta + 1
            lnLinPag = lnLinPag + lnCntRel + 1
            If lnCtaPag Mod 4 = 0 Then
                lsReporte = lsReporte & lsRTFRfm
                lsRTFRfm = ""
            End If
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       If Mid$(lsCodCta, 6, 3) = Producto.gCapAhorros Then
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
          lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
          lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
       Else
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & Chr$(10)
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & Chr$(10)
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
          lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
          lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
       End If
    End If
    lsReporte = lsReporte & lsRTFRfm
End Sub

Private Function RelCta(psCodCta As String, psCodPer As String) As Integer
    Dim lsQrySd1 As String
    Dim rsQrySd1 As ADODB.Recordset
    Set rsQrySd1 = New ADODB.Recordset
    Dim lnCntRel As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    lsQrySd1 = " SELECT PP.cCtaCod, Left(CO.cConsDescripcion,2) cRelaCta, PE.cPersNombre " _
             & " FROM ProductoPersona PP " _
             & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " _
             & " INNER JOIN Constante CO ON PP.nPrdPersRelac = CO.nConsValor And CO.nConsCod = 2005" _
             & " WHERE PP.cCtaCod = '" & psCodCta & "' AND PP.cPersCod <> '" & psCodPer & "' " _
             & " AND PP.nPrdPersRelac = 10"
    
    If lsQrySd1 = "" Then
       Exit Function
    End If
    If rsQrySd1.State = adStateOpen Then rsQrySd1.Close
    
    Set rsQrySd1 = oCon.CargaRecordSet(lsQrySd1)
    
    If rsQrySd1.EOF And rsQrySd1.BOF Then
       rsQrySd1.Close
       Set rsQrySd1 = Nothing
       RelCta = 0
       Exit Function
    Else
       lnCntRel = 0
       With rsQrySd1
            Do While Not .EOF
               Select Case Mid$(psCodCta, 6, 3)
                    Case Producto.gCapAhorros
                         lsRTFRfm = lsRTFRfm & String(26, " ")
                         lnCntRel = lnCntRel + 1
                    Case Producto.gCapPlazoFijo
                         lsRTFRfm = lsRTFRfm & String(55, " ")
                         lnCntRel = lnCntRel + 1
                    Case Producto.gCapCTS
                         lsRTFRfm = lsRTFRfm & String(55, " ")
                         lnCntRel = lnCntRel + 1
               End Select
               lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cPersNombre), False), 34) & Chr(10)
               .MoveNext
            Loop
       End With
       rsQrySd1.Close
       Set rsQrySd1 = Nothing
       RelCta = lnCntRel
    End If

End Function


'***********************************************
' GetInteres: Funcion que devuelve el interes de la cuenta este interes puede ser
' simple o compuesto de acuerdo al ultimo argumento de la funcion
' si este es True es I simple de lo contrario sera compuesto
'***********************************************
Private Function GetInteres(pnTasa As Currency, pnSaldo As Currency, pnDias As Integer, Simple As Boolean) As Double
    Dim lnTasaEfec As Double
    
    If Simple Then
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((lnTasaEfec) * pnDias * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    Else
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((((lnTasaEfec + 1) ^ pnDias) - 1) * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    End If
End Function

Public Function GeneraRepInac(psTipAho As Producto, psTipPer As Integer, psTipMon As Moneda, _
                              psFecpro As String, Optional nTipo As Integer = 1) As String
   Dim lsQrySdo As String
   Dim rsQrySdo As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
        
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And P.cCtaCod Like '___" & lsAgeCod & "%'"
    End If
  
   
    
    oCon.AbreConexion
    lsRTFSdo = ""
    lsMoneda = IIf(psTipMon = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    Select Case psTipPer
        Case PersPersoneria.gPersonaNat
            lsPerson = "PERSONA NATURAL"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurSFL
            lsPerson = "PERSONA JUR. SIN FINES DE LUCRO"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFL
            lsPerson = "PERSONA JUR. CON FINES DE LUCRO"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCMAC
            lsPerson = "CAJAS MUNICIPALES"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCRAC
            lsPerson = "CAJAS RURALES"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLFONCODES
             lsPerson = "FONCODES"
             sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCooperativa
            lsPerson = "COOPERATIVA"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLEdpyme
            lsPerson = "EDPYME"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
    End Select
    
    lsQrySdo = ""
    
        
            If nTipo = 1 Then
                    
                    lsQrySdo = " select p.cctacod,nSdoCnt=p.nsaldo,nsdodis=c.nsaldodisp, "
                    lsQrySdo = lsQrySdo & " tiempoinactiva=datediff(yy,dultcontacto,cast('" & psFecpro & "' as datetime)),Meses=(datediff(mm,dultcontacto,cast('" & psFecpro & "' as datetime)))%12,  "
                    lsQrySdo = lsQrySdo & " CordPag=ca.bordpag, "
                    lsQrySdo = lsQrySdo & " PE.cPersNombre cNomPers, PE.cPersCod cCodPers,a.cagecod,A.CAGEDESCRIPCION "
                    lsQrySdo = lsQrySdo & " from producto p "
                    lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON Pp.cCtaCod = P.cCtaCod "
                    lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON Pe.cPersCod = Pp.cPersCod "
                    lsQrySdo = lsQrySdo & " join captaciones c on c.cctacod=p.cctacod "
                    lsQrySdo = lsQrySdo & " join captacahorros ca on ca.cctacod=p.cctacod "
                    lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                    lsQrySdo = lsQrySdo & " where binactiva=1 and left(p.nprdestado,2) not in ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') and PP.nPrdPersRelac = '" & gCapRelPersTitular & "' "
                    lsQrySdo = lsQrySdo & " and datediff(yy,dultcontacto,getdate())>=1 " & sqlAdd
                    lsQrySdo = lsQrySdo & " and SUBSTRING(P.cCtaCod,9,1)='" & psTipMon & "' and datediff(yy,dultcontacto,getdate())<2 "
                    lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,ca.bOrdPag,p.cctacod"
                    
            ElseIf nTipo = 2 Then
                    
                    lsQrySdo = " select p.cctacod,nSdoCnt=p.nsaldo,nsdodis=c.nsaldodisp, "
                    lsQrySdo = lsQrySdo & " tiempoinactiva=datediff(yy,dultcontacto,cast('" & psFecpro & "' as datetime)),Meses=(datediff(mm,dultcontacto,cast('" & psFecpro & "' as datetime)))%12, "
                    lsQrySdo = lsQrySdo & " CordPag=ca.bordpag, "
                    lsQrySdo = lsQrySdo & " PE.cPersNombre cNomPers, PE.cPersCod cCodPers ,a.cagecod,A.CAGEDESCRIPCION "
                    lsQrySdo = lsQrySdo & " from producto p "
                    lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON Pp.cCtaCod = P.cCtaCod "
                    lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON Pe.cPersCod = Pp.cPersCod "
                    lsQrySdo = lsQrySdo & " join captaciones c on c.cctacod=p.cctacod "
                    lsQrySdo = lsQrySdo & " join captacahorros ca on ca.cctacod=p.cctacod "
                    lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                    lsQrySdo = lsQrySdo & " where binactiva=1 and left(p.nprdestado,2) not in ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') and PP.nPrdPersRelac = '" & gCapRelPersTitular & "'" & sqlAdd
                    lsQrySdo = lsQrySdo & " and SUBSTRING(P.cCtaCod,9,1)='" & psTipMon & "' and datediff(yy,dultcontacto,getdate())>=2"
                    lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,ca.bOrdPag,p.cctacod"
                    
            End If
                  
    
    If lsQrySdo = "" Then
       Exit Function
    End If
    
    If rsQrySdo.State = adStateOpen Then rsQrySdo.Close

    Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
    
    If rsQrySdo.EOF And rsQrySdo.BOF Then
       lsRTFSdo = ""
       GeneraRepInac = lsRTFSdo
    Else
       PrtRepInac rsQrySdo, nTipo
       GeneraRepInac = lsRTFSdo
    End If
End Function

Private Function PrtRepInac(prQryDat As ADODB.Recordset, Optional nTipo As Integer = 1)
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Double, lnTotRet As Double
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Double, lnSdoDis As Double
    Dim lnSdCnFi  As Double, lnSdDiFi As Double
    Dim lnCapita As Double, lsCodCta As String
    Dim lnCapiFi As Double
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnIntMes As Double, lnSdInFi As Double
    Dim lnCntRel As Integer
    Dim lbCtaIde As Boolean
    Dim lnIntCap As Double
    Dim lnIntDev As Double, lnCalInt As Double
    Dim lbCieMes As Boolean
    Dim lnNumDia As Integer, CCodAge As String
    
    If CierreRealizado(2) Then
        lbCieMes = True
    Else
        lbCieMes = False
    End If
    
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE CUENTAS INACTIVAS" & IIf(nTipo = 1, "(DE 1 AÑO Y MENOS DE 2 AÑOS)", "(MAS DE 2 AÑOS)")
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
    lnCntPag = 0
    lbCtaIde = False
    
    With prQryDat
         Do While Not .EOF
             
               lnCarLin = 132
               lsTitRp2 = "AHORROS" & IIf(Not !cOrdPag, "", " CON ORDEN DE PAGO") & " - " & lsPerson
    
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(23, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                  
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                  
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(26, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                  
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
                                   
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                  lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                  lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Or CCodAge <> !cAgeCod Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(23, " ")
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                  
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                  
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFSdo = lsRTFSdo & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1 + " AGENCIA: " + !cAgeDescripcion, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)
               
               
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "    TIEMPO INACTIVA    " & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & " DISPONIBLE  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   "
                  lsRTFSdo = lsRTFSdo & "    (Años)  (Meses)    " & Chr$(10)
                  
               
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10) & Chr$(10)         ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            If !Cctacod = lsCodCta Then
               If lbCtaIde = False Then
                  lnCntRel = RelCta(!Cctacod, !cCodPers)
                  lbCtaIde = True
               End If
            Else
               lsCodCta = !Cctacod
               lbCtaIde = False
               lsRTFSdo = lsRTFSdo & !Cctacod & " "
               If Len(Trim(!cNomPers)) < 50 Then
                  lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43)), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43), " ") & " "
               Else
                  lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43)) & " "
               End If
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               Select Case Mid$(!Cctacod, 6, 3)
                  Case Producto.gCapAhorros
                       lsRTFSdo = lsRTFSdo & JDNum("0.00", 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                  
                       lsRTFSdo = lsRTFSdo & Space(8) & FillText(!tiempoinactiva & Space(8 - Len(!tiempoinactiva)) & !Meses, 25, " ") & Chr(10)
                       
                       lnCapita = lnCapita + 0: lnCapiFi = lnCapiFi + 0
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnCntRel = RelCta(!Cctacod, !cCodPers)
                       lbCtaIde = True
                  
               End Select
            End If
            lnCntCta = lnCntCta + 1
            lnLinPag = lnLinPag + lnCntRel + 1
            
            CCodAge = !cAgeCod
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(23, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
      
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
              
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(26, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
       
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
       
       
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
       lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
       lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
    End If
End Function



Public Function GeneraRepBloq(psTipAho As Producto, psTipPer As Integer, psTipMon As Moneda, _
                              Optional nTipo As Integer = 1) As String
    Dim lsQrySdo As String
    Dim rsQrySdo As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
        
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And P.cCtaCod Like '___" & lsAgeCod & "%'"
    End If
    If psTipAho = gCapCTS Then
     sqlAdd = sqlAdd & " and pb.nblqmotivo in (3,15) "
    ElseIf nTipo = 4 Then
     sqlAdd = sqlAdd & " and pb.nblqmotivo=3 "
    ElseIf nTipo = 5 Then
     sqlAdd = sqlAdd & " and pb.nblqmotivo=15 "
    End If
    
    If nTipo = 2 Then
        sqlAdd = sqlAdd & " and CS.CCODINST='1089800000272' "
    ElseIf nTipo = 3 Then
        sqlAdd = sqlAdd & " and CS.CCODINST<>'1089800000272' "
    End If
    
    
    oCon.AbreConexion
    lsRTFSdo = ""
    lsMoneda = IIf(psTipMon = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    Select Case psTipPer
        Case PersPersoneria.gPersonaNat
            lsPerson = "PERSONA NATURAL"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurSFL
            lsPerson = "PERSONA JUR. SIN FINES DE LUCRO"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFL
            lsPerson = "PERSONA JUR. CON FINES DE LUCRO"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCMAC
            lsPerson = "CAJAS MUNICIPALES"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCRAC
            lsPerson = "CAJAS RURALES"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLFONCODES
             lsPerson = "FONCODES"
             sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLCooperativa
            lsPerson = "COOPERATIVA"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
        Case PersPersoneria.gPersonaJurCFLEdpyme
            lsPerson = "EDPYME"
            sqlAdd = sqlAdd & " and C.nPersoneria = '" & psTipPer & "'"
    End Select
    
    lsQrySdo = ""
    Select Case psTipAho
        Case Producto.gCapAhorros
            
          lsQrySdo = " Select p.cctacod,cordpag=ca.bordpag,cCodPers=PP.CPERSCOD,cNomPers=pe.cpersnombre,nCapital=0,nSdodis=c.nsaldodisp,c.nIntAcum AS nIntDev,nSdoCnt=p.nSaldo,Motivo=pb.cconsdescripcion,a.cagecod,A.CAGEDESCRIPCION "
          lsQrySdo = lsQrySdo & " from producto p "
          lsQrySdo = lsQrySdo & " inner join productopersona pp  on pp.cctacod=p.cctacod "
          lsQrySdo = lsQrySdo & " inner join captaciones c on c.cctacod=p.cctacod "
          lsQrySdo = lsQrySdo & " inner join captacahorros ca on ca.cctacod=c.cctacod "
          lsQrySdo = lsQrySdo & " inner join persona pe on pe.cperscod=pp.cperscod "
          lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(P.CCTACOD,4,2) "
          lsQrySdo = lsQrySdo & " inner join (SELECT distinct p1.nblqmotivo,p1.CCTACOD,p2.cconsdescripcion FROM PRODUCTOBLOQUEOS p1 "
          lsQrySdo = lsQrySdo & " inner join (select nconsvalor,cconsdescripcion from constante where nconscod='2007') p2 on p2.nconsvalor=p1.nblqmotivo "
          lsQrySdo = lsQrySdo & " WHERE  (CMOVNRODBL IS NULL OR CMOVNRODBL='')) PB ON PB.CCTACOD=P.CCTACOD "
          lsQrySdo = lsQrySdo & " WHERE   Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstBloqRetiro, 2) & "','" & Left(CaptacEstado.gCapEstBloqTotal, 2) & "') "
          lsQrySdo = lsQrySdo & " AND PP.nPrdPersRelac in ('" & gCapRelPersTitular & "','" & gCapRelPersRepTitular & "') AND SUBSTRING(P.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & ""
          lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,ca.bordpag,p.cctacod "
                      
        Case Producto.gCapPlazoFijo
        
            lsQrySdo = " Select p.cctacod,cCodPers=PP.CPERSCOD,cNomPers=pe.cpersnombre,nCapital=c.nsaldodisp,c.nIntAcum AS nIntDev,nSdoCont=p.nSaldo,Motivo=pb.cconsdescripcion,a.cagecod,A.CAGEDESCRIPCION "
            lsQrySdo = lsQrySdo & " from producto p "
            lsQrySdo = lsQrySdo & " inner join productopersona pp  on pp.cctacod=p.cctacod "
            lsQrySdo = lsQrySdo & " inner join captaciones c on c.cctacod=p.cctacod "
            lsQrySdo = lsQrySdo & " inner join captacplazofijo cs on cs.cctacod=c.cctacod "
            lsQrySdo = lsQrySdo & " inner join persona pe on pe.cperscod=pp.cperscod "
            lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(P.CCTACOD,4,2) "
            lsQrySdo = lsQrySdo & " inner join (SELECT distinct p1.nblqmotivo,p1.CCTACOD,p2.cconsdescripcion FROM PRODUCTOBLOQUEOS p1 "
            lsQrySdo = lsQrySdo & " inner join (select nconsvalor,cconsdescripcion from constante where nconscod='2007') p2 on p2.nconsvalor=p1.nblqmotivo "
            lsQrySdo = lsQrySdo & " WHERE  (CMOVNRODBL IS NULL OR CMOVNRODBL='')) PB ON PB.CCTACOD=P.CCTACOD "
            lsQrySdo = lsQrySdo & " WHERE   Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstBloqRetiro, 2) & "','" & Left(CaptacEstado.gCapEstBloqTotal, 2) & "') "
            lsQrySdo = lsQrySdo & " AND PP.nPrdPersRelac in ('" & gCapRelPersTitular & "','" & gCapRelPersRepTitular & "') AND SUBSTRING(P.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & ""
            lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,p.cctacod "

                      
        Case Producto.gCapCTS
             If psTipPer = PersPersoneria.gPersonaNat Then
                   lsQrySdo = "Select p.cctacod,cCodPers=PP.CPERSCOD,cNomPers=pe.cpersnombre,nCapital=c.nsaldodisp,nSdoDis=cs.nsaldretiro,nSdoCnt=p.nsaldo,Motivo=pb.cconsdescripcion,a.cagecod,A.CAGEDESCRIPCION  from producto p "
                   lsQrySdo = lsQrySdo & " inner join productopersona pp  on pp.cctacod=p.cctacod "
                   lsQrySdo = lsQrySdo & " inner join captaciones c on c.cctacod=p.cctacod "
                   lsQrySdo = lsQrySdo & " inner join captaccts cs on cs.cctacod=c.cctacod "
                   lsQrySdo = lsQrySdo & " inner join persona pe on pe.cperscod=pp.cperscod "
                   lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(P.CCTACOD,4,2) "
                   lsQrySdo = lsQrySdo & " inner join (SELECT distinct p1.nblqmotivo,p1.CCTACOD,p2.cconsdescripcion FROM PRODUCTOBLOQUEOS p1 "
                   lsQrySdo = lsQrySdo & " inner join (select nconsvalor,cconsdescripcion from constante where nconscod='2007') p2 on p2.nconsvalor=p1.nblqmotivo "
                   lsQrySdo = lsQrySdo & " WHERE  (CMOVNRODBL IS NULL OR CMOVNRODBL='')   ) PB ON PB.CCTACOD=P.CCTACOD "
                   lsQrySdo = lsQrySdo & " WHERE Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstBloqRetiro, 2) & "','" & Left(CaptacEstado.gCapEstBloqTotal, 2) & "') "
                   lsQrySdo = lsQrySdo & " AND C.nPersoneria = '" & psTipPer & "' AND  PP.nPrdPersRelac = '" & gCapRelPersTitular & "' AND SUBSTRING(P.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd & ""
                   lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,p.cctacod "
                            
               
             End If 'RepEstratPlazoFijo
    End Select
    
    If lsQrySdo = "" Then
       Exit Function
    End If
    If rsQrySdo.State = adStateOpen Then rsQrySdo.Close

    Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
    
    If rsQrySdo.EOF And rsQrySdo.BOF Then
       lsRTFSdo = ""
       GeneraRepBloq = lsRTFSdo
    Else
       PrtRepBloqueos rsQrySdo, nTipo
       GeneraRepBloq = lsRTFSdo
    End If


End Function
Private Function PrtRepBloqueos(prQryDat As ADODB.Recordset, Optional nTipo As Integer = 1)
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Double, lnTotRet As Double
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Double, lnSdoDis As Double
    Dim lnSdCnFi  As Double, lnSdDiFi As Double
    Dim lnCapita As Double, lsCodCta As String
    Dim lnCapiFi As Double
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnIntMes As Double, lnSdInFi As Double
    Dim lnCntRel As Integer
    Dim lbCtaIde As Boolean
    Dim lnIntCap As Double
    Dim lnIntDev As Double, lnCalInt As Double
    Dim lbCieMes As Boolean
    Dim lnNumDia As Integer, CCodAge As String
    
    If CierreRealizado(2) Then
        lbCieMes = True
    Else
        lbCieMes = False
    End If
    
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE CUENTAS BLOQUEADAS" & IIf(nTipo = 4, " CREDITOS DIRECTOS", IIf(nTipo = 5, " CREDITOS INDIRECTOS", ""))
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
    lnCntPag = 0
    lbCtaIde = False
    
    With prQryDat
         Do While Not .EOF
         
         
            Select Case Mid$(!Cctacod, 6, 3)
               Case Producto.gCapAhorros
                    lnCarLin = 132
                    lsTitRp2 = "AHORROS" & IIf(Not !cOrdPag, "", " CON ORDEN DE PAGO") & " - " & lsPerson
               Case Producto.gCapPlazoFijo
                    lnCarLin = 119
                    lsTitRp2 = "PLAZO FIJO" & " - " & lsPerson
               Case Producto.gCapCTS
                    lnCarLin = 132
                    lsPerson = "PERSONA NATURAL - " '& IIf(LTrim(RTrim(!cCodInst)) = gsInstCmac, "EMPLEADOS", "CLIENTES")
                    lsTitRp2 = "C T S" & " - " & lsPerson & " " & IIf(nTipo = 2, "EMPLEADOS CMAC - CONVENIO", IIf(nTipo = 3, "OTRAS EMPRESAS - EXTERNOS", ""))
            End Select
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(22, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                     
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(25, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                  lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                  lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Or CCodAge <> !cAgeCod Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(22, " ")
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     If Left(lsCodPro, 7) = "AHORROS" Then
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                     Else
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
                     End If
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFSdo = lsRTFSdo & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1 + " AGENCIA: " + !cAgeDescripcion, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)
               If Mid$(!Cctacod, 6, 3) = Producto.gCapPlazoFijo Then
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   INTERES   "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & " MOTIVO     " & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & "  DEVENGADO  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   " & ""
                  lsRTFSdo = lsRTFSdo & " DE BLOQUEO  " & Chr$(10)
               Else
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   MOTIVO      " & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(20, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & " DISPONIBLE  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   "
                  lsRTFSdo = lsRTFSdo & "   DE BLOQUEO  " & Chr$(10)
                  
               End If
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10) & Chr$(10)         ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            If !Cctacod = lsCodCta Then
               If lbCtaIde = False Then
                  lnCntRel = RelCta(!Cctacod, !cCodPers)
                  lbCtaIde = True
               End If
            Else
               lsCodCta = !Cctacod
               lbCtaIde = False
               lsRTFSdo = lsRTFSdo & !Cctacod & " "
               If Len(Trim(!cNomPers)) < 50 Then
                  lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43)), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43), " ") & " "
               Else
                  lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), IIf(Mid$(!Cctacod, 6, 3) = "233", 42, 43)) & " "
               End If
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               Select Case Mid$(!Cctacod, 6, 3)
                  Case Producto.gCapAhorros
                       lsRTFSdo = lsRTFSdo & JDNum("0.00", 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "

                       lnIntCap = 0
                       'If lbCieMes Then
                       '     lnIntCap = !nSdoDis - !nSdoAnt
                       'Else
                       '     lnIntCap = !nIntMes
                       'End If
                       'lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Space(8) & FillText(!Motivo, 25, " ") & Chr(10)
                       
                       lnCapita = lnCapita + 0: lnCapiFi = lnCapiFi + 0
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnCntRel = RelCta(!Cctacod, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapPlazoFijo
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 12, True, 9, 2) & " "
                       
                       'lnNumDia = DateDiff("d", !dUltMovPF, gdFecSis) - 1
                       
                       'Adicion Miguel 18-03-2004
                       lnCalInt = !nIntDev
                       
                       'lnCalInt = GetInteres(!nTasaInteres, (!nIntDev + !nCapital), lnNumDia, False)
                       
                       lnIntDev = lnCalInt
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCont, 12, True, 9, 2) & "   "
                       lsRTFSdo = lsRTFSdo & FillText(!Motivo, 25, " ") & Chr(10)
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoDis = lnSdoDis + lnIntDev
                       lnSdDiFi = lnSdDiFi + lnIntDev
                       lnSdoCnt = lnSdoCnt + !nSdoCont
                       lnSdCnFi = lnSdCnFi + !nSdoCont
                       lnCntRel = RelCta(!Cctacod, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapCTS
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 11, True, 8, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & "     "
                       
                       lnIntCap = 0
'                       If lbCieMes Then
'                            lnIntCap = !nSdoDis - !nSdoAnt
'                       Else
'                            lnIntCap = !nIntMes
'                       End If
                       lsRTFSdo = lsRTFSdo & FillText(!Motivo, 25, " ") & Chr(10)
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnCntRel = RelCta(!Cctacod, !cCodPers)
                       lbCtaIde = True
               End Select
            End If
            lnCntCta = lnCntCta + 1
            lnLinPag = lnLinPag + lnCntRel + 1
            
            CCodAge = !cAgeCod
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(22, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
       
       Else
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & Chr$(10)
          
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(25, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & Chr$(10)
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
       lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
       lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
    End If
End Function
Public Function GeneraSaldos(psTipAho As Producto, psTipPer As Integer, psTipMon As Moneda, _
                             psFchPro As Date, Optional nTipo As Integer = 1) As String
    Dim lsQrySdo As String
    Dim rsQrySdo As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
        
    If lsAgeCod = "" Then
        sqlAdd = ""
    ElseIf lsAgeCod <> "X4" And lsAgeCod <> "X2" Then
        sqlAdd = " And PR.cCtaCod Like '___" & lsAgeCod & "%'"
    End If
    
    If nTipo = 2 Then
        sqlAdd = sqlAdd & " and CAPCTS.CCODINST='1089800000272' "
    ElseIf nTipo = 3 Then
        sqlAdd = sqlAdd & " and CAPCTS.CCODINST<>'1089800000272' "
    ElseIf nTipo = 6 Then
          sqlAdd = sqlAdd & " and CAPCTS.CCODINST='1089800000272' and pr.nprdestado not in (1100,1200) and pr.cctacod not in (select cctacod from productobloqueos where nblqmotivo in (12,3) and cmovnrodbl is null and substring(cctacod,6,3)='" & Producto.gCapCTS & "') "
    End If
    

    
    oCon.AbreConexion
    lsRTFSdo = ""
    lsMoneda = IIf(psTipMon = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    
    If nTipo <> 5 Then
        Select Case psTipPer
            Case PersPersoneria.gPersonaNat
                 lsPerson = "PERSONA NATURAL"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurSFL
                 lsPerson = "PERSONA JUR. SIN FINES DE LUCRO"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFL
                 lsPerson = "PERSONA JUR. CON FINES DE LUCRO"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFLCMAC
                 lsPerson = "CAJAS MUNICIPALES"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFLCRAC
                 lsPerson = "CAJAS RURALES"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFLFONCODES
                 lsPerson = "FONCODES"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFLCooperativa
                 lsPerson = "COOPERATIVA"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
            Case PersPersoneria.gPersonaJurCFLEdpyme
                 lsPerson = "EDPYME"
                 sqlAdd = sqlAdd & " AND CAP.nPersoneria = '" & psTipPer & "'"
        End Select
   End If
    
    
    If nTipo = 5 Then
        lsPerson = "PERS. NATURAL-JUR. SIN FINES LUCRO"
    End If
    
    lsQrySdo = ""
    Select Case psTipAho
         Case Producto.gCapAhorros
            If nTipo = 5 Then
                        
                        lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAPAHO.bOrdPag cOrdPag, CAP.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt,"
                        lsQrySdo = lsQrySdo & "  dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers, "
                        lsQrySdo = lsQrySdo & "  CAPAHO.dUltContacto AS dUltMovAC, PR.nTasaInteres nTasaIntAC, CAPAHO.nSaldoAnterior AS nSdoAnt, "
                        lsQrySdo = lsQrySdo & " a.cagecod,A.CAGEDESCRIPCION,CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacAhorros CAPAHO ON PR.cCtaCod = CAPAHO.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null "
                        lsQrySdo = lsQrySdo & "  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & " AND CAP.nPersoneria IN ('" & PersPersoneria.gPersonaNat & "','" & PersPersoneria.gPersonaJurSFL & "') AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "'" & sqlAdd
                        lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,CAPAHO.bOrdPag, PR.cCtaCod "
        
            Else
               
                   If lsAgeCod <> "X4" And lsAgeCod <> "X2" Then
                        lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAPAHO.bOrdPag cOrdPag, CAP.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt,"
                        lsQrySdo = lsQrySdo & "  dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers, "
                        lsQrySdo = lsQrySdo & "  CAPAHO.dUltContacto AS dUltMovAC, PR.nTasaInteres nTasaIntAC, CAPAHO.nSaldoAnterior AS nSdoAnt, "
                        lsQrySdo = lsQrySdo & " a.cagecod,A.CAGEDESCRIPCION,CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacAhorros CAPAHO ON PR.cCtaCod = CAPAHO.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null "
                        lsQrySdo = lsQrySdo & "  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "'" & sqlAdd
                        lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,CAPAHO.bOrdPag, PR.cCtaCod "
                   ElseIf lsAgeCod = "X4" Then
                        lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAPAHO.bOrdPag cOrdPag, CAP.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt,"
                        lsQrySdo = lsQrySdo & "  dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers, "
                        lsQrySdo = lsQrySdo & "  CAPAHO.dUltContacto AS dUltMovAC, PR.nTasaInteres nTasaIntAC, CAPAHO.nSaldoAnterior AS nSdoAnt, "
                        lsQrySdo = lsQrySdo & " cagecod='X4',CAGEDESCRIPCION='NASCA - PALPA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacAhorros CAPAHO ON PR.cCtaCod = CAPAHO.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        'lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null "
                        lsQrySdo = lsQrySdo & "  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' and substring(PR.cCtaCod,4,2) in ('04','08')" & sqlAdd
                        lsQrySdo = lsQrySdo & " ORDER BY CAPAHO.bOrdPag, PR.cCtaCod "
                   ElseIf lsAgeCod = "X2" Then
                        lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAPAHO.bOrdPag cOrdPag, CAP.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt,"
                        lsQrySdo = lsQrySdo & "  dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers, "
                        lsQrySdo = lsQrySdo & "  CAPAHO.dUltContacto AS dUltMovAC, PR.nTasaInteres nTasaIntAC, CAPAHO.nSaldoAnterior AS nSdoAnt, "
                        lsQrySdo = lsQrySdo & " cagecod='X2',CAGEDESCRIPCION='CAÑETE - MALA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacAhorros CAPAHO ON PR.cCtaCod = CAPAHO.cCtaCod "
                        lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        'lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null "
                        lsQrySdo = lsQrySdo & "  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "'and substring(PR.cCtaCod,4,2) in ('02','07') " & sqlAdd
                        lsQrySdo = lsQrySdo & " ORDER BY CAPAHO.bOrdPag, PR.cCtaCod "
                    End If
            End If

        Case Producto.gCapPlazoFijo
             
             If nTipo = 5 Then
                lsQrySdo = lsQrySdo & " SELECT PR.cCtaCod cCodCta, CAP.nSaldoDisp AS nCapital, CAP.nIntAcum AS nIntDev,  PR.nSaldo AS nSdoCnt, "
                lsQrySdo = lsQrySdo & "  PE.cPersNombre cNomPers, PP.cPersCod cCodPers, PR.nTasaInteres, CAPPF.dAuxiliar AS dUltMovPF , "
                lsQrySdo = lsQrySdo & "  a.cagecod,A.CAGEDESCRIPCION,CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ),NPLAZO "
                lsQrySdo = lsQrySdo & "  FROM Producto PR "
                lsQrySdo = lsQrySdo & "  INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                lsQrySdo = lsQrySdo & "  INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                lsQrySdo = lsQrySdo & "  INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod "
                lsQrySdo = lsQrySdo & "  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                lsQrySdo = lsQrySdo & "  inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                lsQrySdo = lsQrySdo & "  INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & "  LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                lsQrySdo = lsQrySdo & "  WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                lsQrySdo = lsQrySdo & "  AND CAP.nPersoneria IN ('" & PersPersoneria.gPersonaNat & "','" & PersPersoneria.gPersonaJurSFL & "')  AND  PP.nPrdPersRelac in (10,12) AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "'" & sqlAdd
                lsQrySdo = lsQrySdo & "  ORDER BY a.cagecod,PR.cCtaCod ,PP.nPrdPersRelac "
              Else
                    If lsAgeCod <> "X4" And lsAgeCod <> "X2" Then
                        lsQrySdo = lsQrySdo & " SELECT PR.cCtaCod cCodCta, CAP.nSaldoDisp AS nCapital, CAP.nIntAcum AS nIntDev,  PR.nSaldo AS nSdoCnt, "
                        lsQrySdo = lsQrySdo & "  PE.cPersNombre cNomPers, PP.cPersCod cCodPers, PR.nTasaInteres, CAPPF.dAuxiliar AS dUltMovPF , "
                        lsQrySdo = lsQrySdo & "  a.cagecod,A.CAGEDESCRIPCION,CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ),NPLAZO "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & "  INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        lsQrySdo = lsQrySdo & "  inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & "  INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & "  WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & "  AND  PP.nPrdPersRelac in (10,12) AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd
                        lsQrySdo = lsQrySdo & "  ORDER BY a.cagecod,PR.cCtaCod  ,PP.nPrdPersRelac "
                    ElseIf lsAgeCod = "X4" Then
                        lsQrySdo = lsQrySdo & " SELECT PR.cCtaCod cCodCta, CAP.nSaldoDisp AS nCapital, CAP.nIntAcum AS nIntDev,  PR.nSaldo AS nSdoCnt, "
                        lsQrySdo = lsQrySdo & "  PE.cPersNombre cNomPers, PP.cPersCod cCodPers, PR.nTasaInteres, CAPPF.dAuxiliar AS dUltMovPF , "
                        lsQrySdo = lsQrySdo & "  cagecod='X4',CAGEDESCRIPCION='NASCA - PALPA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ),NPLAZO "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & "  INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        'lsQrySdo = lsQrySdo & "  inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & "  INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & "  WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & "  AND  PP.nPrdPersRelac in (10,12) AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' and substring(PR.cCtaCod,4,2) in ('04','08')" & sqlAdd
                        lsQrySdo = lsQrySdo & "  ORDER BY PR.cCtaCod ,PP.nPrdPersRelac "
                    ElseIf lsAgeCod = "X2" Then
                        lsQrySdo = lsQrySdo & " SELECT PR.cCtaCod cCodCta, CAP.nSaldoDisp AS nCapital, CAP.nIntAcum AS nIntDev,  PR.nSaldo AS nSdoCnt, "
                        lsQrySdo = lsQrySdo & "  PE.cPersNombre cNomPers, PP.cPersCod cCodPers, PR.nTasaInteres, CAPPF.dAuxiliar AS dUltMovPF , "
                        lsQrySdo = lsQrySdo & "  cagecod='X2',CAGEDESCRIPCION='CAÑETE - MALA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ),NPLAZO "
                        lsQrySdo = lsQrySdo & "  FROM Producto PR "
                        lsQrySdo = lsQrySdo & "  INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod "
                        lsQrySdo = lsQrySdo & "  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                        'lsQrySdo = lsQrySdo & "  inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                        lsQrySdo = lsQrySdo & "  INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                        lsQrySdo = lsQrySdo & "  Where E2.cNroDoc Is Null  GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                        lsQrySdo = lsQrySdo & "  WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
                        lsQrySdo = lsQrySdo & "  AND  PP.nPrdPersRelac in (10,12) AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' and substring(PR.cCtaCod,4,2) in ('02','07') " & sqlAdd
                        lsQrySdo = lsQrySdo & "  ORDER BY PR.cCtaCod,PP.nPrdPersRelac "
                    End If
              End If
                
                      
        Case Producto.gCapCTS
             If psTipPer = PersPersoneria.gPersonaNat Then
                
             If lsAgeCod <> "X4" And lsAgeCod <> "X2" Then
                lsQrySdo = "  SELECT CAPCTS.cCtaCod cCodCta, CAPCTS.nSaldRetiro As nSdoDis, PR.nSaldo nSdoCnt,  CAP.nSaldoDisp As nCapital,dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, "
                lsQrySdo = lsQrySdo & " case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end AS dUltMovCTS, CAPCTS.cCodInst, PE.cPersNombre cNomPers, PP.cPersCod cCodPers, "
                lsQrySdo = lsQrySdo & " a.cagecod,A.CAGEDESCRIPCION,CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                lsQrySdo = lsQrySdo & " FROM Producto PR "
                lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN CaptacCTS CAPCTS ON PR.cCtaCod = CAPCTS.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                lsQrySdo = lsQrySdo & " LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " Where E2.cNroDoc Is Null "
                lsQrySdo = lsQrySdo & " GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') "
                lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' " & sqlAdd
                lsQrySdo = lsQrySdo & " ORDER BY a.cagecod,CAPCTS.cCodInst, PR.cCtaCod "
             ElseIf lsAgeCod = "X4" Then
                lsQrySdo = "  SELECT CAPCTS.cCtaCod cCodCta, CAPCTS.nSaldRetiro As nSdoDis, PR.nSaldo nSdoCnt,  CAP.nSaldoDisp As nCapital,dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, "
                lsQrySdo = lsQrySdo & " case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end AS dUltMovCTS, CAPCTS.cCodInst, PE.cPersNombre cNomPers, PP.cPersCod cCodPers, "
                lsQrySdo = lsQrySdo & " cagecod='X4',CAGEDESCRIPCION='NASCA - PALPA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                lsQrySdo = lsQrySdo & " FROM Producto PR "
                lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN CaptacCTS CAPCTS ON PR.cCtaCod = CAPCTS.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                'lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                lsQrySdo = lsQrySdo & " LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " Where E2.cNroDoc Is Null "
                lsQrySdo = lsQrySdo & " GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') "
                lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' and substring(PR.cCtaCod,4,2) in ('04','08') " & sqlAdd
                lsQrySdo = lsQrySdo & " ORDER BY CAPCTS.cCodInst, PR.cCtaCod "
             ElseIf lsAgeCod = "X2" Then
                lsQrySdo = "  SELECT CAPCTS.cCtaCod cCodCta, CAPCTS.nSaldRetiro As nSdoDis, PR.nSaldo nSdoCnt,  CAP.nSaldoDisp As nCapital,dbo.GetInteresGanado(PR.cCtaCod,'" & Format(psFchPro, gsFormatoFecha) & "','S') AS nIntMes, "
                lsQrySdo = lsQrySdo & " case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end AS dUltMovCTS, CAPCTS.cCodInst, PE.cPersNombre cNomPers, PP.cPersCod cCodPers, "
                lsQrySdo = lsQrySdo & " cagecod='X2',CAGEDESCRIPCION='CAÑETE - MALA',CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ) "
                lsQrySdo = lsQrySdo & " FROM Producto PR "
                lsQrySdo = lsQrySdo & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN CaptacCTS CAPCTS ON PR.cCtaCod = CAPCTS.cCtaCod "
                lsQrySdo = lsQrySdo & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod "
                'lsQrySdo = lsQrySdo & " inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2) "
                lsQrySdo = lsQrySdo & " LEFT JOIN (select MONTO=SUM(NMONTO),CCTACOD from DOCRECCAPTA D "
                lsQrySdo = lsQrySdo & " INNER JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=1 AND CIFCTA IS NOT NULL) E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " LEFT JOIN (SELECT CNRODOC,CIFCTA FROM DOCRECEST WHERE NESTADO=2 AND CIFCTA IS NOT NULL ) E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA "
                lsQrySdo = lsQrySdo & " Where E2.cNroDoc Is Null "
                lsQrySdo = lsQrySdo & " GROUP BY D.CCTACOD) L ON L.CCTACOD=PP.CCTACOD "
                lsQrySdo = lsQrySdo & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') "
                lsQrySdo = lsQrySdo & " AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "'  and substring(PR.cCtaCod,4,2) in ('02','07') " & sqlAdd
                lsQrySdo = lsQrySdo & " ORDER BY CAPCTS.cCodInst, PR.cCtaCod "
                End If
             End If 'RepEstratPlazoFijo
    End Select
    
    If lsQrySdo = "" Then
       Exit Function
    End If
    If rsQrySdo.State = adStateOpen Then rsQrySdo.Close

    Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
    
    If rsQrySdo.EOF And rsQrySdo.BOF Then
       lsRTFSdo = ""
       GeneraSaldos = lsRTFSdo
    ElseIf psTipAho = gCapAhorros Then
       PrtSaldos rsQrySdo, nTipo
       GeneraSaldos = lsRTFSdo
    ElseIf psTipAho = gCapPlazoFijo Then
       PrtSaldosTI rsQrySdo, nTipo
       GeneraSaldos = lsRTFSdo
    ElseIf psTipAho = gCapCTS Then
        PrtSaldos rsQrySdo, nTipo
       GeneraSaldos = lsRTFSdo
    End If
End Function



Private Function PrtSaldosTI(prQryDat As ADODB.Recordset, Optional nTipo As Integer = 1)
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Double, lnTotRet As Double
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Double, lnSdoDis As Double
    Dim lnCapita As Double, lsCodCta As String
    Dim lnCapiFi As Double, lnSdDiFi As Double, lnSdCnFi As Double
    Dim lnCheEnV As Double, lnCheEnVFi As Double
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnIntMes As Double, lnSdInFi As Double
    Dim lnCntRel As Integer
    Dim lbCtaIde As Boolean
    Dim lnIntCap As Double
    Dim lnIntDev As Double, lnCalInt As Double
    Dim lbCieMes As Boolean
    Dim lnNumDia As Integer, CCodAge As String
    Dim lsCuenta As String
    'Dim codigocta As String
    If CierreRealizado(2) Then
        lbCieMes = True
    Else
        lbCieMes = False
    End If
    
    'codigocta = ""
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE SALDO DE CUENTA(S)"
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
    lnCntPag = 0
    lbCtaIde = False
    
    With prQryDat
         Do While Not .EOF
         
         
            Select Case Mid$(!cCodCta, 6, 3)
               Case Producto.gCapAhorros
                    lnCarLin = 132
                    lsTitRp2 = "AHORROS" & IIf(Not !cOrdPag, "", " CON ORDEN DE PAGO") & " - " & lsPerson
               Case Producto.gCapPlazoFijo
                    lnCarLin = 119
                    lsTitRp2 = "PLAZO FIJO" & " - " & lsPerson
         '      Case Producto.gCapCTS
          '          lnCarLin = 132
          '          lsPerson = "PERSONA NATURAL - " & IIf(LTrim(RTrim(!cCodInst)) = gsInstCmac, "EMPLEADOS", "CLIENTES")
          '          lsTitRp2 = "C T S" & " - " & lsPerson & " " & IIf(ntipo = 2, "EMPLEADOS CMAC - CONVENIO", IIf(ntipo = 3, "OTRAS EMPRESAS - EXTERNOS", ""))
            End Select
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                  
                  
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 11, True, 9, 2) & ""
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                     
                  Else
                  
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 11, True, 9, 2) & ""
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 11, True, 9, 2) & ""
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
                     
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(17, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 11, True, 9, 2) & ""
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 11, True, 9, 2) & ""
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & ""
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & ""
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 11, True, 9, 2) & "" '& Chr$(10)
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & Chr$(10)
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                  lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
                  lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Or CCodAge <> !cAgeCod Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 11, True, 9, 2) & ""
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 11, True, 9, 2) & ""
                     If Left(lsCodPro, 7) = "AHORROS" Then
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                     Else
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 11, True, 9, 2) & ""  '& Chr$(10)
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
                     End If
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFSdo = lsRTFSdo & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1 + " AGENCIA: " + !cAgeDescripcion, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)
               If Mid$(!cCodCta, 6, 3) = Producto.gCapPlazoFijo Then
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "           "
                  lsRTFSdo = lsRTFSdo & "  INTERES  "
                  lsRTFSdo = lsRTFSdo & "   SALDO   " '& " "
                  lsRTFSdo = lsRTFSdo & "  CHEQUES  "
                  lsRTFSdo = lsRTFSdo & "       " '& " "
                  lsRTFSdo = lsRTFSdo & "  ULTIMO  " & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "  CAPITAL  "
                  lsRTFSdo = lsRTFSdo & "  DEVENG.  "
                  lsRTFSdo = lsRTFSdo & "   CONT.   " '& " "
                  lsRTFSdo = lsRTFSdo & "EN VALORIZ."
                  lsRTFSdo = lsRTFSdo & " PLAZO " '& " "
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & Chr$(10)
               Else
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "              "
                  lsRTFSdo = lsRTFSdo & "    SALDO     "
                  lsRTFSdo = lsRTFSdo & "    SALDO     "
                  lsRTFSdo = lsRTFSdo & "   CHEQUES    "
                  lsRTFSdo = lsRTFSdo & "   INTERESES  " & " "
                  lsRTFSdo = lsRTFSdo & "  ULTIMO" & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & " DISPONIBLE  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   "
                  lsRTFSdo = lsRTFSdo & " EN VALORIZ. "
                  lsRTFSdo = lsRTFSdo & "    DEL  MES" & " "
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & Chr$(10)
               End If
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10) & Chr$(10)         ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            'If !ccodcta = lsCodCta Then
            '   If lbCtaIde = False Then
            '      lnCntRel = RelCta(!ccodcta, !cCodPers)
                  '********AGREGADOS MPBR*******
                   'lsRTFSdo = String(19, " ")
                   'If Len(Trim(!cNomPers)) < 50 Then
                    '    lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37), 37, " ") & " "
                   'Else
                   '     lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37) & " "
                   'End If
                  'lbCtaIde = False
                  '********AGREGADOS MPBR*********
            '      lbCtaIde = True
            '   End If
           ' Else
             If !cCodCta <> lsCodCta Then
                lsRTFSdo = lsRTFSdo & !cCodCta & " "
             Else
                lsRTFSdo = lsRTFSdo & String(19, " ")
             End If
               
               lbCtaIde = False
               
               If Len(Trim(!cNomPers)) < 50 Then
                  lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37), 37, " ") & "" & IIf(!cCodCta = lsCodCta, Chr(10), "")
               Else
                  lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37) & "" & IIf(!cCodCta = lsCodCta, Chr(10), "")
               End If
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               
             If !cCodCta <> lsCodCta Then
                Select Case Mid$(!cCodCta, 6, 3)
                       Case Producto.gCapAhorros
                            lsRTFSdo = lsRTFSdo & JDNum("0.00", 12, True, 9, 2) & " "
                            lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                            lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                            lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "

                            lnIntCap = 0
                            If lbCieMes Then
                                 lnIntCap = !nSdoDis - !nSdoAnt
                            Else
                                 lnIntCap = !nIntMes
                            End If
                            lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                            lsRTFSdo = lsRTFSdo & Format(!dUltMovAC, gsFormatoFechaView) & Chr(10)
                            
                            lnCapita = lnCapita + 0: lnCapiFi = lnCapiFi + 0
                            lnSdoCnt = lnSdoCnt + !nSdoCnt
                            lnSdoDis = lnSdoDis + !nSdoDis
                            lnIntMes = lnIntMes + lnIntCap
                            lnSdCnFi = lnSdCnFi + !nSdoCnt
                            lnCheEnV = lnCheEnV + !chequesenv
                            lnCheEnVFi = lnCheEnVFi + !chequesenv
                            lnSdDiFi = lnSdDiFi + !nSdoDis
                            lnSdInFi = lnSdInFi + lnIntCap
                            lnCntRel = RelCta(!cCodCta, !cCodPers)
                            lbCtaIde = True
                        Case Producto.gCapPlazoFijo
                            lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 11, True, 9, 2) & ""
                            
                            lnNumDia = DateDiff("d", Format(!dUltMovPF, "dd/mm/yyyy"), gdFecSis) - 1
                            
                            'Adicion Miguel 18-03-2004
                            lnCalInt = !nIntDev
                            
                            'lnCalInt = GetInteres(!nTasaInteres, (!nIntDev + !nCapital), lnNumDia, False)
                            
                            lnIntDev = lnCalInt
                            lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntDev)), 11, True, 9, 2) & ""
                            lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 11, True, 9, 2) & ""
                            lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 11, True, 9, 2) & ""
                            lsRTFSdo = lsRTFSdo & JDNum(!nPlazo, 7, True, 5, 0) & " "
                            lsRTFSdo = lsRTFSdo & Format(!dUltMovPF, gsFormatoFechaView) & Chr(10)
                            
                            lnCapita = lnCapita + !nCapital
                            lnCapiFi = lnCapiFi + !nCapital
                            lnSdoDis = lnSdoDis + lnIntDev
                            lnSdDiFi = lnSdDiFi + lnIntDev
                            lnSdoCnt = lnSdoCnt + !nSdoCnt
                            lnSdCnFi = lnSdCnFi + !nSdoCnt
                            lnCheEnV = lnCheEnV + !chequesenv
                            lnCheEnVFi = lnCheEnVFi + !chequesenv
                            lnCntRel = RelCta(!cCodCta, !cCodPers)
                            lbCtaIde = True
                         Case Producto.gCapCTS
                           lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 11, True, 8, 2) & " "
                           lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                           lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                           lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "
                           lnIntCap = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "
    '                       If lbCieMes Then
    '                            lnIntCap = !nSdoDis - !nSdoAnt
    '                       Else
    '                            lnIntCap = !nIntMes
    '                       End If
                           lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                           lsRTFSdo = lsRTFSdo & Format(!dUltMovCTS, gsFormatoFechaView) & Chr(10)
                           
                           lnCapita = lnCapita + !nCapital
                           lnCapiFi = lnCapiFi + !nCapital
                           lnSdoCnt = lnSdoCnt + !nSdoCnt
                           lnSdCnFi = lnSdCnFi + !nSdoCnt
                           lnCheEnV = lnCheEnV + !chequesenv
                           lnCheEnVFi = lnCheEnVFi + !chequesenv
                           lnSdoDis = lnSdoDis + !nSdoDis
                           lnSdDiFi = lnSdDiFi + !nSdoDis
                           lnIntMes = lnIntMes + lnIntCap
                           lnSdInFi = lnSdInFi + lnIntCap
                           lnCntRel = RelCta(!cCodCta, !cCodPers)
                           lbCtaIde = True
                       End Select
             End If
                lnCntCta = lnCntCta + 1
                lnLinPag = lnLinPag + lnCntRel + 1
                lsCuenta = Mid(!cCodCta, 6, 3)
            
                lsCodCta = !cCodCta
                CCodAge = !cAgeCod
            'codigocta = !ccodcta
                .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ") & IIf(lsCuenta = Producto.gCapCTS, "", " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 11, True, 9, 2) & ""
       If Left(lsCodPro, 7) = "AHORROS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & ""
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & ""
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & ""
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 11, True, 9, 2) & ""
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 11, True, 9, 2) & "" '& Chr$(10)
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(17, " ") & IIf(lsCuenta = Producto.gCapCTS, "", " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 11, True, 9, 2) & ""
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 11, True, 9, 2) & ""
       If Left(lsCodPro, 7) = "AHORROS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 11, True, 9, 2) & "" '& Chr$(10)
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & Chr$(10)
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
       lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
       lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
    End If
End Function

Private Function PrtSaldos(prQryDat As ADODB.Recordset, Optional nTipo As Integer = 1)
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Double, lnTotRet As Double
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Double, lnSdoDis As Double
    Dim lnCapita As Double, lsCodCta As String
    Dim lnCapiFi As Double, lnSdDiFi As Double, lnSdCnFi As Double
    Dim lnCheEnV As Double, lnCheEnVFi As Double
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnIntMes As Double, lnSdInFi As Double
    Dim lnCntRel As Integer
    Dim lbCtaIde As Boolean
    Dim lnIntCap As Double
    Dim lnIntDev As Double, lnCalInt As Double
    Dim lbCieMes As Boolean
    Dim lnNumDia As Integer, CCodAge As String
    Dim lsCuenta As String
    Dim codigocta As String
    If CierreRealizado(2) Then
        lbCieMes = True
    Else
        lbCieMes = False
    End If
    
    codigocta = ""
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE SALDO DE CUENTA(S)"
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
    lnCntPag = 0
    lbCtaIde = False
    
    With prQryDat
         Do While Not .EOF
         
         
            Select Case Mid$(!cCodCta, 6, 3)
               Case Producto.gCapAhorros
                    lnCarLin = 132
                    lsTitRp2 = "AHORROS" & IIf(Not !cOrdPag, "", " CON ORDEN DE PAGO") & " - " & lsPerson
               Case Producto.gCapPlazoFijo
                    lnCarLin = 119
                    lsTitRp2 = "PLAZO FIJO" & " - " & lsPerson
               Case Producto.gCapCTS
                    lnCarLin = 132
                    lsPerson = "PERSONA NATURAL - " & IIf(LTrim(RTrim(!cCodInst)) = gsInstCmac, "EMPLEADOS", "CLIENTES")
                    lsTitRp2 = "CTS" & " - " & lsPerson & " " & IIf(nTipo = 2, "EMPLEADOS CMAC - CONVENIO", IIf(nTipo = 3, "OTRAS EMPRESAS - EXTERNOS", ""))
            End Select
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(16, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                  
                  
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                     
                  ElseIf Left(lsCodPro, 3) = "CTS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                     
                  Else
                  
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
                     
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(19, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " " '& Chr$(10)
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & Chr$(10)
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
                  lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
                  lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Or CCodAge <> !cAgeCod Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
                     lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(16, " ") & IIf(Mid$(!cCodCta, 6, 3) = Producto.gCapCTS, "", " ")
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     If Left(lsCodPro, 7) = "AHORROS" Then
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                     ElseIf Left(lsCodPro, 3) = "CTS" Then
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "  '& Chr$(10)
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
                      
                     Else
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "  '& Chr$(10)
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
                     End If
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(12)             ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFSdo = lsRTFSdo & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1 + " AGENCIA: " + !cAgeDescripcion, lsTitRp2, lsMoneda, lsNumPag) & Chr$(10)
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)
               If Mid$(!cCodCta, 6, 3) = Producto.gCapPlazoFijo Then
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   INTERES   "
                  lsRTFSdo = lsRTFSdo & "   SALDO    " '& " "
                  lsRTFSdo = lsRTFSdo & "    CHEQUES  "
                  lsRTFSdo = lsRTFSdo & "    ULTIMO" & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & "  DEVENGADO  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE  " '& " "
                  lsRTFSdo = lsRTFSdo & "   EN VALORIZ."
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & Chr$(10)
               Else
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "  CHEQUES    "
                  lsRTFSdo = lsRTFSdo & "   INTERESES" & " "
                  lsRTFSdo = lsRTFSdo & "  ULTIMO" & Chr$(10)
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(14, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & " DISPONIBLE  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   "
                  lsRTFSdo = lsRTFSdo & " EN VALORIZ. "
                  lsRTFSdo = lsRTFSdo & "    DEL  MES" & " "
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & Chr$(10)
               End If
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10) & Chr$(10)         ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            If !cCodCta = lsCodCta Then
               If lbCtaIde = False Then
                  lnCntRel = RelCta(!cCodCta, !cCodPers)
                  '********AGREGADOS MPBR*******
                   'lsRTFSdo = String(19, " ")
                   'If Len(Trim(!cNomPers)) < 50 Then
                    '    lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37), 37, " ") & " "
                   'Else
                   '     lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37) & " "
                   'End If
                  'lbCtaIde = False
                  '********AGREGADOS MPBR*********
                  lbCtaIde = True
               End If
            Else
               lsCodCta = !cCodCta
               lbCtaIde = False
               lsRTFSdo = lsRTFSdo & !cCodCta & " "
               If Len(Trim(!cNomPers)) < 50 Then
                  lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37), 37, " ") & " "
               Else
                  lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 37) & " "
               End If
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               Select Case Mid$(!cCodCta, 6, 3)
                  Case Producto.gCapAhorros
                       lsRTFSdo = lsRTFSdo & JDNum("0.00", 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "

                       lnIntCap = 0
                       If lbCieMes Then
                            lnIntCap = !nSdoDis - !nSdoAnt
                       Else
                            lnIntCap = !nIntMes
                       End If
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovAC, gsFormatoFechaView) & Chr(10)
                       
                       lnCapita = lnCapita + 0: lnCapiFi = lnCapiFi + 0
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnIntMes = lnIntMes + lnIntCap
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnCheEnV = lnCheEnV + !chequesenv
                       lnCheEnVFi = lnCheEnVFi + !chequesenv
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnSdInFi = lnSdInFi + lnIntCap
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapPlazoFijo
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 12, True, 9, 2) & " "
                       
                       lnNumDia = DateDiff("d", !dUltMovPF, gdFecSis) - 1
                       
                       'Adicion Miguel 18-03-2004
                       lnCalInt = !nIntDev
                       
                       'lnCalInt = GetInteres(!nTasaInteres, (!nIntDev + !nCapital), lnNumDia, False)
                       
                       lnIntDev = lnCalInt
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovPF, gsFormatoFechaView) & Chr(10)
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoDis = lnSdoDis + lnIntDev
                       lnSdDiFi = lnSdDiFi + lnIntDev
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnCheEnV = lnCheEnV + !chequesenv
                       lnCheEnVFi = lnCheEnVFi + !chequesenv
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapCTS
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 11, True, 8, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!chequesenv, 12, True, 9, 2) & " "
                       lnIntCap = !nIntMes
'                       If lbCieMes Then
'                            lnIntCap = !nSdoDis - !nSdoAnt
'                       Else
'                            lnIntCap = !nIntMes
'                       End If
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovCTS, gsFormatoFechaView) & Chr(10)
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnCheEnV = lnCheEnV + !chequesenv
                       lnCheEnVFi = lnCheEnVFi + !chequesenv
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnIntMes = lnIntMes + lnIntCap
                       lnSdInFi = lnSdInFi + lnIntCap
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
               End Select
            End If
            lnCntCta = lnCntCta + 1
            lnLinPag = lnLinPag + lnCntRel + 1
            lsCuenta = Mid(!cCodCta, 6, 3)
               
            CCodAge = !cAgeCod
            codigocta = !cCodCta
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(16, " ") & IIf(lsCuenta = Producto.gCapCTS, "", " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
       ElseIf Left(lsCodPro, 3) = "CTS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & Chr$(10)
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " " '& Chr$(10)
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnV)), 12, True, 9, 2) & Chr$(10)
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & Chr$(10)             ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(19, " ") & IIf(lsCuenta = Producto.gCapCTS, "", " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
       ElseIf Left(lsCodPro, 3) = "CTS" Then
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & Chr$(10)
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " " '& Chr$(10)
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCheEnVFi)), 12, True, 9, 2) & Chr$(10)
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & Chr$(12)             ' hasta 160 caracteres x línea
       lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0: lnCheEnV = 0
       lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0: lnCheEnVFi = 0
    End If
End Function

Public Function ReporteNoCobroInactivas(Optional ByVal CCodAge As String = "") As String
 Dim oCon As DConecta, ssql As String, rstemp As ADODB.Recordset, sqlaux As String
 
 sqlaux = ""
 
 Set oCon = New DConecta
 
 If CCodAge <> "" Then
    sqlaux = " and substring(p.cctacod,4,2)='" & CCodAge & "'"
 End If
    
     oCon.AbreConexion
     
        ssql = "  Select p.Cctacod, "
        ssql = ssql & " cpersnombre=( Select top  1  f.cpersnombre from persona f join  productopersona pp  on pp.cperscod=f.cperscod    where pp.nprdpersrelac=10 and pp.cctacod=p.cctacod ), "
        ssql = ssql & "  SaldoCnt=p.nsaldo,SaldoDisp=c.nsaldodisp "
        ssql = ssql & " from Producto p "
        ssql = ssql & " join Captaciones c on c.cctacod=p.cctacod "
        ssql = ssql & "  join CaptacAhorros ca on ca.cctacod=p.cctacod "
        ssql = ssql & " where p.nprdestado not in ('1400','1300' ) and ca.bnocobroinactivas=1 " & sqlaux
        ssql = ssql & "    ORDER BY substring(p.cctacod,4,2),CPERSNOMBRE  "
        
        
        Set rstemp = oCon.CargaRecordSet(ssql)
            
     oCon.CierraConexion
     
     
    Dim lnCarLin As Integer, i As Integer
    Dim lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsCodAge As String
    Dim lsCuenta As String
    Dim clsage As DAgencias
    Set clsage = New DAgencias
    
    
    lsNumPag = ""
    'lnLinPag = 65
    lnCarLin = 95
    'lsTitRp1 = " REPORTE DE CUENTAS EXONERADAS DE COBRO DE INACTIVAS " & IIf(CcodAge <> "", clsage.NombreAgencia(CcodAge), "")
    lnLinPag = 0
    lsCodAge = ""
    
    lsRTFSdo = "" '& Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    i = i + 1
    While Not rstemp.EOF
    
         If lsCodAge = "" Or lnLinPag > 65 Then
            
            lsRTFSdo = lsRTFSdo & Chr(10) & Chr(10) & Chr(10) & Chr(10)
            lsCodAge = CCodAge
            lsRTFSdo = lsRTFSdo & Space(10) & "CMAC ICA" & Space(70) & "FECHA: " & gdFecSis & Chr(10)
            lsRTFSdo = lsRTFSdo & Space(10) & Space(8) & Space(70) & "HORA:" & GetHoraServer() & Chr(10) & Chr(10)
            
            lsRTFSdo = lsRTFSdo & Space(15) & " REPORTE DE CUENTAS EXONERADAS DE COBRO DE INACTIVAS " & IIf(CCodAge <> "", clsage.NombreAgencia(CCodAge), "") & Chr(10) & Chr(10)
            
            
            lsRTFSdo = lsRTFSdo & Space(10) & String(lnCarLin, "-") & Chr$(10)
            lsRTFSdo = lsRTFSdo & Space(10) & "   NRO.   "
            lsRTFSdo = lsRTFSdo & "   CUENTA           "
            lsRTFSdo = lsRTFSdo & "   TITULAR                           "
            lsRTFSdo = lsRTFSdo & "   SALDO CONT. "
            lsRTFSdo = lsRTFSdo & "   SALDO DISP. " & Chr(10)
            lsRTFSdo = lsRTFSdo & Space(10) & String(lnCarLin, "-") & Chr$(10)
                        
         End If
                
            lsRTFSdo = lsRTFSdo & Space(8) & FillNum(Trim(Str(i)), 10, " ") & Space(1) & Trim(rstemp!Cctacod) & Space(1) & FillText(Left(PstaNombre(ImpreCarEsp(rstemp!cPersNombre), False), 37), 37, " ") & JDNum(Trim(Str(rstemp!SaldoCnt)), 11, True, 9, 2) & Space(3) & JDNum(Trim(Str(rstemp!SaldoDisp)), 11, True, 9, 2) & Chr(10)
             
            If lnLinPag = 65 Then
                lsRTFSdo = lsRTFSdo & Chr$(12)
            End If
             
            lnLinPag = lnLinPag + 1
            i = i + 1
            
             rstemp.MoveNext
    Wend
     
    ReporteNoCobroInactivas = lsRTFSdo
     
    Set rstemp = Nothing
    Set oCon = Nothing
        
End Function
Public Function CUPONIMPRIMIR(ByVal CNUMSORTEO As String, ByVal cNombre As String, ByVal cNumRangoIni As String, ByVal cNumRangoFin As String, ByVal cNumCuenta As String, ByVal cTipo As Integer, Optional ByVal CnumCupon As String, Optional ByVal cNumRangoIniPrint As String, Optional ByVal cNumRangoFinPrint As String) As String
'clsage.NombreAgencia(CCodAge)
  Dim clsage As DAgencias, CCodAge As String
  Set clsage = New DAgencias
  
   CCodAge = Mid(cNumCuenta, 4, 2)
  
Dim nFicSal As Integer, sFecha As String, sHora As String
Dim sSep As Integer, sIni As Integer, sMax As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsCadArg As String
Dim sCad As String

ETIQ:

On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)


sCad = sCad & Chr$(27) & Chr$(64)

sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5


lsNomAge = sNomAge

'cNumRangoIni As String, ByVal cNumRangoFin

'scad = scad & Chr$(10);
If cTipo = "1" Then
    sCad = sCad & lsNegritaOn 'Activa Negrita
    sCad = sCad & Chr(10) & Chr(10) & Chr(10) & lsNegritaOn 'Activa Negrita
    sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & " " & clsage.NombreAgencia(CCodAge) & Chr(10) & Chr(10)
    sCad = sCad & Space(sIni) & UCase(cNombre) & Chr(10)
    sCad = sCad & "DEL CUPON NRO. " & cNumRangoIni & " AL " & cNumRangoFin & Chr(10)
    sCad = sCad & "NRO. CUENTA " & cNumCuenta & " " & Chr(10)
    sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & " " & clsage.NombreAgencia(CCodAge) & Chr(10) & Chr(10) & Chr(12)
    sCad = sCad & lsNegritaOff 'Desactiva Negrita
        
ElseIf cTipo = "2" Then
    For i = Val(cNumRangoIniPrint) To Val(cNumRangoFinPrint)
            sCad = sCad & lsNegritaOn 'Activa Negrita
            sCad = sCad & Chr(10) & Chr(10) & Chr(10) & lsNegritaOn 'Activa Negrita
            sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CStr(i) & Chr(10) & Chr(10)
            sCad = sCad & Space(sIni) & UCase(cNombre) & Chr(10)
            sCad = sCad & "DEL CUPON NRO. " & cNumRangoIni & " AL " & cNumRangoFin & Chr(10)
            sCad = sCad & "RANGO IMPRESO DEL " & cNumRangoIniPrint & " AL " & cNumRangoFinPrint & Chr(10)
            sCad = sCad & "NRO. " & CnumCupon & Chr(10)
            sCad = sCad & "NRO. CUENTA " & cNumCuenta & " " & Chr(10)
            sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & Chr(10) & Chr(10) & Chr(12)
            sCad = sCad & lsNegritaOff 'Desactiva Negrita
        
    Next i
ElseIf cTipo = "3" Then
    
            sCad = sCad & lsNegritaOn 'Activa Negrita
            sCad = sCad & Chr(10) & Chr(10) & Chr(10) & lsNegritaOn 'Activa Negrita
            sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CStr(i) & Chr(10) & Chr(10)
            sCad = sCad & Space(sIni) & UCase(cNombre) & Chr(10)
            sCad = sCad & "DEL CUPON NRO. " & cNumRangoIni & " AL " & cNumRangoFin & Chr(10)
            sCad = sCad & "NRO. " & CnumCupon & Chr(10)
            sCad = sCad & "NRO. CUENTA " & cNumCuenta & " " & Chr(10)
            sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & Chr(10) & Chr(10) & Chr(12)
            sCad = sCad & lsNegritaOff 'Desactiva Negrita

End If


Exit Function
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If


End Function
Public Function CONSOLIDADOSORTEO()

End Function
Public Function CUPONESBATCH(ByVal CNUMSORTEO As String, ByVal CCodAge As String, Optional CcodAgeAnt As String) As String
  Dim oCon As DConecta, ssql As String, rstemp As Recordset
  Dim clsage As DAgencias
   Set clsage = New DAgencias
   
    Dim nFicSal As Integer, sFecha As String, sHora As String, sSep As Integer
    Dim sIni As Integer, lsNegritaOn As String, lsNegritaOff As String
    Dim lsNroExt As String, lnTope As Integer
    Dim sCad As String
  
  Set oCon = New DConecta
  Set rstemp = New Recordset
  
     oCon.AbreConexion
      
    ssql = " Select c.cperscod,cNumCta=c.cctacod ,cnombre=p.cpersnombre,cNumRangoIni=c.nrangoini,cNumRangoFin=c.nrangofin from cuentasorteo c "
    ssql = ssql & " join persona p on p.cperscod=c.cperscod  WHERE left(CNUMSORTEO,6)='" & CNUMSORTEO & "' and  substring(c.cctacod,4,2)='" & CCodAge & "'"
    ssql = ssql & " and (bcancelar<>1 or bcancelar is null )  and (  banular<>1 or banular is null) "
    ssql = ssql & " order by p.cpersnombre  "
    
    Set rstemp = oCon.CargaRecordSet(ssql)
      
     oCon.CierraConexion
     Set oCon = Nothing
     
    
ETIQ:
    
    On Error GoTo Error
    
    lnTope = 0 '6 'Tope de lineas en Boleta
    
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
   
    
    sCad = sCad & Chr$(27) & Chr$(64)
    
    sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
    sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita
    
    sSep = 15
    sIni = 1

    
    
    
    'lsNomAge = sNomAge
    
    'cNumRangoIni As String, ByVal cNumRangoFin
    
    'scad = scad & Chr$(10);
    While Not rstemp.EOF
        sCad = sCad & lsNegritaOn 'Activa Negrita
        sCad = sCad & Chr(10) & Chr(10) & Chr(10) & lsNegritaOn 'Activa Negrita
        sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & " " & clsage.NombreAgencia(CCodAge) & Chr(10) & Chr(10)
        sCad = sCad & Space(sIni) & UCase(rstemp!cNombre) & Chr(10)
        sCad = sCad & "DEL CUPON NRO. " & rstemp!cNumRangoIni & " AL " & rstemp!cNumRangoFin & Chr(10)
        sCad = sCad & "NRO. CUENTA " & rstemp!cNumCta & " " & Chr(10)
        sCad = sCad & Space(sIni) & "CMACICA - PREMIAHORRO " & CNUMSORTEO & " " & clsage.NombreAgencia(CCodAge) & Chr(10) & Chr(10) & Chr(12)
        sCad = sCad & lsNegritaOff 'Desactiva Negrita
        
        rstemp.MoveNext
    Wend
           
           CUPONESBATCH = sCad
           
    Exit Function
Error:
        Close nFicSal
        If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
            GoTo ETIQ
        End If
      
      
End Function

Public Function PLANILLACONSOLSORTEO(ByVal CNUMSORTEO As String, Optional ByVal CCodAge As String)
Dim oCon As DConecta, ssql As String, rstemp As Recordset, sqlaux As String, sNumSorteo As String
 
 sqlaux = ""
 
 Set oCon = New DConecta
 
    
 Set rstemp = New Recordset
    
     oCon.AbreConexion
     
        ssql = "  Select   c.cnumsorteo, c.cctacod, c.nsaldo, c.cperscod,p.cpersnombre,"
        ssql = ssql & " C.NRANGOINI , C.NRANGOFIN, C.NNUMTICKETS, C.cMovNro, bGanador, nNumganador, bCancelar, bAnular "
        ssql = ssql & " From cuentasorteo C join persona p on p.cperscod=c.cperscod "
        ssql = ssql & " where (bcancelar<>1 or bcancelar is null )  and (  banular<>1 or banular is null)    and substring(cctacod,4,2)='" & CCodAge & "'"
        ssql = ssql & " order by p.cpersnombre "
        
        
        Set rstemp = oCon.CargaRecordSet(ssql)
            
     oCon.CierraConexion
     
     
    Dim lnCarLin As Integer, i As Integer
    Dim lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsCodAge As String
    Dim lsCuenta As String
    Dim clsage As DAgencias
    Set clsage = New DAgencias
    
    
    lsNumPag = ""
    'lnLinPag = 65
    lnCarLin = 120
    'lsTitRp1 = " REPORTE DE CUENTAS EXONERADAS DE COBRO DE INACTIVAS " & IIf(CcodAge <> "", clsage.NombreAgencia(CcodAge), "")
    lnLinPag = 0
    lsCodAge = ""
    
    lsRTFSdo = "" '& Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    i = i + 1
    
    
    While Not rstemp.EOF
                
         
          
         If lsCodAge = "" Or lnLinPag > 65 Then
            If CcodAgeAnt <> Mid(rstemp!Cctacod, 4, 2) And CcodAgeAnt <> "" Then
                lsRTFSdo = lsRTFSdo & Chr$(12)
                CcodAgeAnt = Mid(rstemp!Cctacod, 4, 2)
            End If
           
            
            lsRTFSdo = lsRTFSdo & Chr(10) & Chr(10) & Chr(10) & Chr(10)
            lsRTFSdo = lsRTFSdo & Space(10) & "CMAC ICA" & Space(70) & "FECHA: " & gdFecSis & Chr(10)
            lsRTFSdo = lsRTFSdo & Space(10) & Space(8) & Space(70) & "HORA:" & GetHoraServer() & Chr(10) & Chr(10)
            
            lsRTFSdo = lsRTFSdo & Space(15) & " REPORTE CONSOLIDADO DE SORTEO  " & CNUMSORTEO & IIf(CCodAge <> "", " - " & clsage.NombreAgencia(CCodAge), "") & Chr(10) & Chr(10)
            
            
            lsRTFSdo = lsRTFSdo & Space(8) & String(lnCarLin, "-") & Chr$(10)
            lsRTFSdo = lsRTFSdo & Space(8) & "  NRO.   "
            lsRTFSdo = lsRTFSdo & "   CUENTA           "
            lsRTFSdo = lsRTFSdo & "   TITULAR                           "
            lsRTFSdo = lsRTFSdo & "   SALDO    "
            lsRTFSdo = lsRTFSdo & "  RANGO INICIAL "
            lsRTFSdo = lsRTFSdo & "  RANGO FINAL "
            lsRTFSdo = lsRTFSdo & "  NRO CUPONES " & Chr$(10)
            lsRTFSdo = lsRTFSdo & Space(8) & String(lnCarLin, "-") & Chr$(10)
                 lnLinPag = 0
         End If
                
            lsRTFSdo = lsRTFSdo & Space(6) & FillNum(Trim(Str(i)), 8, " ") & Space(1) & Trim(rstemp!Cctacod) & Space(1) & FillText(Left(PstaNombre(ImpreCarEsp(rstemp!cPersNombre), False), 37), 37, " ") & JDNum(Trim(Str(rstemp!nSaldo)), 11, True, 9, 2) & Space(3) & JDNum(Trim(Str(rstemp!NRANGOINI)), 11, True, 11, 0) & Space(3) & JDNum(Trim(Str(rstemp!NRANGOFIN)), 11, True, 10, 0) & Space(3) & JDNum(Trim(Str(rstemp!NNUMTICKETS)), 11, True, 10, 0) & Chr(10)
             
             
            
            'If lsCodAge = "04" Then Stop
             
            If lnLinPag = 65 Then
                lsRTFSdo = lsRTFSdo & Chr$(12)
            End If
             
            lsCodAge = Mid(rstemp!Cctacod, 4, 2)
            
            lnLinPag = lnLinPag + 1
            i = i + 1
                        
                       
            rstemp.MoveNext
             
    Wend
     
  PLANILLACONSOLSORTEO = lsRTFSdo
     
    Set rstemp = Nothing
    Set oCon = Nothing
        

End Function

Public Function RepEstratPlazoFijo() As String
    Dim lsReporte As String
    Dim rsPlazo As ADODB.Recordset
    Set rsPlazo = New ADODB.Recordset
    Dim nPlazos As Integer
    Dim J As Integer
    Dim nRangos(1 To 6, 1 To 2) As Integer
    Dim sRangos(1 To 6) As String
    Dim nNum As Long, nNumCTS As Long
    Dim nMonto As Currency, nMontoCTS As Currency
    Dim lsNum As String, lsMonto As String
    Dim nTotalNum As Long
    Dim nTotalMonto As Currency
    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsNumPag As String
    Dim lnLinPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim VSQL As String
    
    oCon.AbreConexion
        
    nPlazos = 6
    
    nRangos(1, 1) = 0:      nRangos(1, 2) = 30
    nRangos(2, 1) = 31:     nRangos(2, 2) = 60
    nRangos(3, 1) = 61:     nRangos(3, 2) = 90
    nRangos(4, 1) = 91:     nRangos(4, 2) = 120
    nRangos(5, 1) = 121:    nRangos(5, 2) = 360
    nRangos(6, 1) = 361:    nRangos(6, 2) = 9999
    
    sRangos(1) = "Menor a 30  "
    sRangos(2) = "31 a 60     "
    sRangos(3) = "61 a 90     "
    sRangos(4) = "91 a 120    "
    sRangos(5) = "121 a 360   "
    sRangos(6) = "Mayor a 360 "
    
    lsNumPag = "   1"
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE PLAZO FIJO POR VENCIMIENTO"
    lsTitRp2 = ""
    lsReporte = CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, "", lsNumPag) & Chr$(10)
    lsReporte = lsReporte & String(lnCarLin, "-") & Chr$(10)              ' hasta 160 caracteres x línea
    
    For J = 1 To 2
        nTotalNum = 0
        nTotalMonto = 0
        If J = 1 Then
            lsReporte = lsReporte & "MONEDA NACIONAL" & Chr$(10)
        Else
            lsReporte = lsReporte & "MONEDA EXTRANJERA" & Chr$(10)
        End If
        lsReporte = lsReporte & String(17, "-") & Chr$(10)
        lsReporte = lsReporte & "PLAZO" & Space(20) & "NUMERO" & Space(18) & "MONTO" & Chr$(10)
        For i = 1 To nPlazos
            nNum = 0
            nMonto = 0
            VSQL = " Select Count(PR.cCtaCod) Num, Sum(nSaldo) Monto" _
                 & " From CaptacPlazoFijo PF " _
                 & " Inner Join Producto PR On PF.cCtaCod = PR.cCtaCod Where " _
                 & " nPlazo between " & nRangos(i, 1) & " and " & nRangos(i, 2) & " and " _
                 & " Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                 & " And Substring(PR.cCtaCod,9,1) = '" & J & "'  And Left(PR.cCtaCod,5) Like '___" & lsAgeCod & "'"
            If rsPlazo.State = adStateOpen Then
                rsPlazo.Close
            End If
            
            Set rsPlazo = oCon.CargaRecordSet(VSQL)
            
            If rsPlazo.EOF And rsPlazo.BOF Then
                nNum = 0
                nMonto = 0
            Else
                nNum = IIf(IsNull(rsPlazo!Num), 0, rsPlazo!Num)
                nMonto = IIf(IsNull(rsPlazo!Monto), 0, rsPlazo!Monto)
            End If
            
            If i = 6 Then
                GetSaldoCTS J, nNumCTS, nMontoCTS
                nNum = nNum + nNumCTS
                nMonto = nMonto + nMontoCTS
            End If
            nTotalNum = nTotalNum + nNum
            nTotalMonto = nTotalMonto + nMonto
            lsNum = Str(nNum)
            lsMonto = Format$(nMonto, "#,##0.00")
            lsReporte = lsReporte & sRangos(i) & Space(14) & JDNum(lsNum, 5, False, 5, 0) & Space(10) & JDNum(lsMonto, 13, True, 10, 2) & Chr$(10)
        Next i
        lsReporte = lsReporte & "TOTAL       " & Space(14) & JDNum(Str(nTotalNum), 5, False, 5, 0) & Space(10) & JDNum(Str(nTotalMonto), 13, True, 10, 2) & Chr$(10)
        lsReporte = lsReporte & Chr$(10)
    Next J
    Set rsPlazo = Nothing
    RepEstratPlazoFijo = lsReporte
End Function

Public Function RepPlazoFijovencimiento(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pnMoneda As Moneda) As String
    Dim lsCadena As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim SQL As String
    Dim lsCodCtaAnt As String
    Dim lsCodCta As String * 18
    Dim lsPersona As String * 40
    Dim lsApertura As String * 10
    Dim lsRelacion As String * 15
    Dim lsPlazo As String * 5
    Dim lsItem As String * 5
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    Dim lnContador As Integer
    Dim lsMonto As String * 15
    Dim lnMonto As Currency
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cap.cctacod like '___" & lsAgeCod & "%'"
    End If
    
    SQL = " Select CAP.cCtaCod, cPersNombre, dApertura, nPlazo, cConsDescripcion, PRD.nSaldo From Captaciones CAP" _
        & " Inner Join CaptacPlazoFijo CPF ON CAP.cCtaCod = CPF.cCtaCod" _
        & " Inner Join ProductoPersona PP ON CAP.cCtaCod = PP.cCtaCod" _
        & " Inner Join Producto PRD On PRD.cCtaCod = CAP.cCtaCod" _
        & " Inner Join Persona PE ON PP.cPersCod = PE.cPersCod" _
        & " Inner Join Constante CO On  PP.nPrdPersRelac = CO.nConsValor And nConsCod = 2005" _
        & " Where dRenovacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'" _
        & " And Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Substring(cap.cctacod,9,1) = " & pnMoneda & " " & sqlAdd & " order by cap.cctacod"

    oCon.AbreConexion
           
    Set rs = oCon.CargaRecordSet(SQL)
    
    lsCadena = ""
    lsCodCtaAnt = ""
    lnContador = 0
    lnMonto = 0
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("VENCIMIENTO DE PLAZOS FIJOS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
        lsCadena = lsCadena & Encabezado("Item;6; ;3;Cuenta;10; ;10;Apertura;8; ;2;Plazo;8;Saldo;15;Persona;20; ;23;Relacion;10; ;3;", lnItem)
        
        While Not rs.EOF
            If lsCodCtaAnt = rs!Cctacod Then
                lsCodCta = ""
                lsPersona = rs!cPersNombre
                lsApertura = ""
                lsRelacion = rs!cConsDescripcion
                lsPlazo = ""
                lsItem = ""
                lsMonto = ""
            Else
                lnContador = lnContador + 1
                lsCodCta = rs!Cctacod
                lsPersona = rs!cPersNombre
                lsApertura = Format(rs!dApertura, gsFormatoFechaView)
                lsRelacion = rs!cConsDescripcion
                RSet lsPlazo = rs!nPlazo
                lsItem = Format(lnContador, "00000")
                RSet lsMonto = Format(rs!nSaldo, "#,##0.00")
                lnMonto = lnMonto + rs!nSaldo
            End If
            lnItem = lnItem + 1
            
            lsCadena = lsCadena & "  " & lsItem & "  " & lsCodCta & "  " & lsApertura & "  " & lsPlazo & "  " & lsMonto & "  " & lsPersona & "  " & lsRelacion & Chr(10)
            
            If lnItem > 54 Then
                lsCadena = lsCadena & CabeceraPagina("VENCIMIENTO DE PLAZOS FIJOS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
                lsCadena = lsCadena & Encabezado("Item;6; ;3;Cuenta;10; ;10;Apertura;8; ;2;Plazo;8;Saldo;15;Persona;20; ;23;Relacion;10; ;3;", lnItem)
            End If
            
            lsCodCtaAnt = rs!Cctacod
            rs.MoveNext
        Wend
        
        RSet lsMonto = Format(lnMonto, "#,##0.00")
        lsCadena = lsCadena & Encabezado(" ;47;" & lsMonto & ";16; ;55;", lnItem)
        
    End If
            
    RepPlazoFijovencimiento = lsCadena
End Function


Private Sub GetSaldoCTS(ByVal psMoneda As Moneda, ByRef pnNum As Long, ByRef pnMonto As Currency)
    Dim rsCTS As ADODB.Recordset
    Set rsCTS = New ADODB.Recordset
    Dim nNum As Long
    Dim nMonto As Currency
    Dim VSQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    VSQL = " Select Count(PR.cCtaCod) Num, Sum(nSaldo) Monto from Producto PR" _
         & " Inner Join CaptacCTS CTS On PR.cCtaCod = CTS.cCtaCod" _
         & " where Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
         & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "' And Left(PR.cCtaCod,5) Like '___" & lsAgeCod & "'"
    
    If rsCTS.State = adStateOpen Then
        rsCTS.Close
    End If
    
    Set rsCTS = oCon.CargaRecordSet(VSQL)
    
    If rsCTS.EOF And rsCTS.BOF Then
        pnNum = 0
        pnMonto = 0
    Else
        pnNum = IIf(IsNull(rsCTS!Num), 0, rsCTS!Num)
        pnMonto = IIf(IsNull(rsCTS!Monto), 0, rsCTS!Monto)
    End If
End Sub

Public Function RepCtaIna() As String
    Dim lsCad As String
    lsCad = ""
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS SOLES - CON ORDEN PAGO -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, 1, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS SOLES - SIN ORDEN PAGO -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, 0, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS DOLARES - CON ORDEN PAGO -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, 1, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS DOLARES - SIN ORDEN PAGO -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, 0, True)
    RepCtaIna = lsCad
End Function

Public Function RepCtaInmobilizada(pdFI As Date, pdFF As Date) As String
    Dim lsCad As String
    lsCad = ""
    lsCad = lsCad & GetCtaInmobilizadaAho("AHORRO INMOVILIZADA SOLES -", Moneda.gMonedaNacional, pdFI, pdFF, 0, 0, 1, True)
    lsCad = lsCad & GetCtaInmobilizadaAho("AHORRO INMOVILIZADA DOLARES -", Moneda.gMonedaExtranjera, pdFI, pdFF, 0, 0, 1, True)
    RepCtaInmobilizada = lsCad
End Function


Public Function GetCtaInaAho(psTit As String, psMoneda As Moneda, psProd As String, lnPagina As Long, lnItem As Long, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As ADODB.Recordset
    Set rsAho = New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 55
    Dim lnMonto  As Currency
    Dim lsMonto  As String * 12
    Dim lsFecha As String * 10
    Dim lsUltCnt As String * 10
    Dim lsContador As String * 4
    Dim lnAcum As Currency
    Dim lnAcumItem As Integer
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnContador As Integer
    
    lsCadena = ""
    sqlAho = " Select PR.cCtaCod Cuenta, PE.cPersNombre Nombre, PR.nSaldo Monto," _
           & " CAP.dApertura FechaApe, CAPAHO.dUltContacto dUltCntAC" _
           & " FROM Producto PR" _
           & " INNER JOIN CaptacAhorros CAPAHO On PR.cCtaCod = CAPAHO.cCtaCod" _
           & " INNER JOIN Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
           & " INNER JOIN ProductoPersona PP On PP.cCtaCod = PR.cCtaCod" _
           & " INNER JOIN Persona PE On PP.cPersCod = PE.cPersCod" _
           & " WHERE bInactiva = 1 and bOrdPag = " & psOrdPag & " and SUBSTRING(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
           & " And Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
           & " ORDER BY PR.cCtaCod"
    
    oCon.AbreConexion
    Set rsAho = oCon.CargaRecordSet(sqlAho)
             
    If Not (rsAho.EOF And rsAho.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        lsCadena = lsCadena & Encabezado("ITEM;4;CUENTA;20;NOMBRE;25; ;32;SALDO;12; ;2;FEC.APE;10; ;2;ULT.CNT;10; ;1;", lnItem)
        
        lsCodCtaAnt = ""
        lnContador = 0
        lnAcum = 0
        While Not rsAho.EOF
            If rsAho!Cuenta <> lsCodCtaAnt Then
                lnContador = lnContador + 1
                lsContador = FillNum(Str(lnContador), 4, "0")
                lnAcum = lnAcum + rsAho!Monto
                lnMonto = rsAho!Monto
                lsFecha = Format$(rsAho!FechaApe, gsFormatoFechaView)
                lsUltCnt = Format$(rsAho!dUltcntAC, gsFormatoFechaView)
            
            Else
                lsContador = ""
                lnMonto = 0
                lsFecha = ""
                lsUltCnt = ""
            End If
            lsCodCta = IIf(rsAho!Cuenta = lsCodCtaAnt, "", rsAho!Cuenta)
            lsNomPers = PstaNombre(rsAho!Nombre)
            lsMonto = IIf(lnMonto = 0, "", JDNum(Str(lnMonto), 12, True, 9, 2))
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(2) & lsMonto & Space(2) & lsFecha & Space(2) & lsUltCnt & Chr(10)
            lnAcumItem = lnAcumItem + 1
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("ITEM;4;CUENTA;20;NOMBRE;25; ;32;SALDO;12; ;2;FEC.APE;10; ;2;ULT.CNT;10; ;1;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & Chr$(10)
        lsCadena = lsCadena & String(119, "=") & Chr$(10)
        lsCadena = lsCadena & lsContador & Space(20) & "TOTAL SALDO" & Space(46) & JDNum(Trim(Str(lnAcum)), 14, True, 11, 2) & Chr$(10)
        lsCadena = lsCadena & String(119, "=") & Chr$(10)
        lsCadena = lsCadena & Chr(12)
    End If
    
    GetCtaInaAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function

Public Function GetCtaInmobilizadaAho(psTit As String, psMoneda As Moneda, pdIni As Date, pdFin As Date, lnPagina As Long, lnItem As Long, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As ADODB.Recordset
    Set rsAho = New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lnMonto  As Currency
    Dim lsMonto  As String * 12
    Dim lsFechaI As String * 10
    Dim lsFechaF As String * 10
    Dim lsContador As String * 4
    Dim lnAcum As Currency
    Dim lnAcumItem As Integer
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnContador As Integer
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(MC.cCtaCod,4,2)  = '" & lsAgeCod & "'"
    End If
        
        
    lsCadena = ""
    sqlAho = " Select MC.cCtaCod, PRD.nSaldo, 0 Interes, dbo.FechaMov(M.cMovNro) FFin, " _
           & " (Select dbo.FechaMov(Max(MA.cMovNro)) From Mov MA" _
           & " Inner Join MovCap MCA On MA.nMovNro = MCA.nMovNro" _
           & " Where MA.cOpeCod = '200803' And MA.nMovFlag = " & MovFlag.gMovFlagVigente & " And MCA.cCtaCod = MC.cCtaCod) FIni" _
           & " From Mov M" _
           & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
           & " Inner Join Producto PRD On PRD.cCtaCod = MC.cCtaCod" _
           & " Where M.cOpeCod = '200406'" _
           & " And Left(M.cMovNro,8) Between '" & Format(pdIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, pdFin), gsFormatoMovFecha) & "' And M.nMovFlag = " & MovFlag.gMovFlagVigente & "" _
           & " And Substring(MC.cCtaCod,9,1) = '" & psMoneda & "' " & sqlAdd
    
    oCon.AbreConexion
    Set rsAho = oCon.CargaRecordSet(sqlAho)
             
    If Not (rsAho.EOF And rsAho.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTit & " " & sqlAdd, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        lsCadena = lsCadena & Encabezado("Item;4; ;1;Inm.;10; ;5;Cancel;10; ;5;cCodCta;10; ;10;Saldo;10; ;5;Int.Ult.M;11;", lnItem)
        
        lnContador = 0
        lnAcum = 0
        While Not rsAho.EOF
            lnContador = lnContador + 1
            lsContador = Format(lnContador, "0000")
            lnAcum = lnAcum + rsAho!nSaldo
            lnAcumItem = lnAcumItem + 1
            
            lsCodCta = rsAho!Cctacod
            RSet lsMonto = Format(rsAho!nSaldo, "#,##0.00")
            lsFechaI = rsAho!FIni
            lsFechaF = rsAho!FFin
            
            lsCadena = lsCadena & lsContador & " " & lsFechaI & " " & lsFechaF & " " & lsCodCta & " " & lsMonto & " " & lsContador & Chr(10)
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("Inm.;10; ;5;Cancel;10; ;5;cCodCta;10; ;10;Saldo;10; ;5;Int.Ult.M;11;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!cCodCta
            rsAho.MoveNext
        Wend
        
        lsCadena = lsCadena & Chr$(10)
        lsCadena = lsCadena & String(119, "=") & Chr$(10)
        lsCadena = lsCadena & lsContador & Space(20) & "TOTAL SALDO" & Space(50) & JDNum(Trim(Str(lnAcum)), 14, True, 11, 2) & Chr$(10)
        lsCadena = lsCadena & String(119, "=") & Chr$(10)
        lsCadena = lsCadena & Chr(12)
    End If
    
    GetCtaInmobilizadaAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function


Function ConsolidadoInactivas(dFecIni As Date, dFecFin As Date) As String
    Dim Sql1 As String
    Dim Sql2 As String
    Dim Sql3 As String
    Dim rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim Rs3 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim lsCad As String
    Dim lnNum1 As Long
    Dim lnMon1 As Currency
    Dim lnNum2 As Long
    Dim lnMon2 As Currency
    Dim lnNum3 As Long
    Dim lnMon3 As Currency
    Dim lnTotNumero As Long
    Dim lnTotMonto As Currency
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    lsCad = lsCad & Chr(10)
    lsCad = lsCad & Space(10) & "CAJA MUNICIPAL DE " & Space(50) & ArmaFecha(ldFecSis) & Chr(10)
    lsCad = lsCad & Space(10) & "    TRUJILLO      " & Space(50) & Str(Time()) & Chr(10)
    lsCad = lsCad & Space(10) & Space(30) & "RESUMEN DE DESCUENTO DE INACTIVAS" & Chr(10)
    lsCad = lsCad & Space(10) & Space(30) & "---------------------------------" & Chr(10) & Chr(10)
    lsCad = lsCad & Space(10) & "Agencia : " & lsNomAge & Chr(10) & Chr(10)
    For i = 1 To 2
        lnNum1 = 0: lnMon1 = 0: lnNum2 = 0: lnMon2 = 0: lnNum3 = 0
        lnMon3 = 0: lnTotNumero = 0: lnTotMonto = 0
        Sql1 = " SELECT COUNT(MC.nMovNro) as Numero ,ISNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoEstActInac & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & i & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
             & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
        
        Sql2 = " SELECT COUNT(MC.nMovNro) as Numero , IsNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoDctoInactiva & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & i & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
             & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
             
        Sql3 = " SELECT COUNT(MC.nMovNro) as Numero , IsNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoCancInact & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & i & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
             & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
        
        Set rs1 = oCon.CargaRecordSet(Sql1)
        Set Rs2 = oCon.CargaRecordSet(Sql2)
        Set Rs3 = oCon.CargaRecordSet(Sql3)
        
        lnNum1 = IIf(rs1.EOF And rs1.BOF, 0, rs1!Numero)
        lnMon1 = IIf(rs1.EOF And rs1.BOF, 0, rs1!Monto)
        lnNum2 = IIf(Rs2.EOF And Rs2.BOF, 0, Rs2!Numero)
        lnMon2 = IIf(Rs2.EOF And Rs2.BOF, 0, Rs2!Monto)
        lnNum3 = IIf(Rs3.EOF And Rs3.BOF, 0, Rs3!Numero)
        lnMon3 = IIf(Rs3.EOF And Rs3.BOF, 0, Rs3!Monto)
        rs1.Close
        Rs2.Close
        Rs3.Close
        Set rs1 = Nothing
        Set Rs2 = Nothing
        Set Rs3 = Nothing
        lnTotNumero = lnNum2 + lnNum3
        lnTotMonto = lnMon2 + lnMon3
        
    lsCad = lsCad & Space(10) & "Moneda : " & IIf(i = 1, "SOLES", "DOLARES") & Chr(10)
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
    lsCad = lsCad & Space(10) & "Descripción                   " & Space(2) & "Nro     " & Space(2) & "     Monto     " & Chr(10)
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
    lsCad = lsCad & Space(10) & "Paso de Activas a Inactivas : " & Space(2) & JDNum(Str(lnNum1), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon1), 13, True, 10, 2) & Chr(10)
    lsCad = lsCad & Space(10) & "Descuento de Inactivas      : " & Space(2) & JDNum(Str(lnNum2), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon2), 13, True, 10, 2) & Chr(10)
    lsCad = lsCad & Space(10) & "Cancelación de Inactivas    : " & Space(2) & JDNum(Str(lnNum3), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon3), 13, True, 10, 2) & Chr(10) & Chr(10)
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
    lsCad = lsCad & Space(10) & "                      TOTAL : " & Space(2) & JDNum(Str(lnTotNumero), 6, False, 6, 0) & Space(2) & JDNum(Str(lnTotMonto), 13, True, 10, 2) & Chr(10)
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10) & Chr(10)
    Next
    ConsolidadoInactivas = lsCad
End Function

Public Function RepCtaApe(pdFecIni As Date, pdFecFin As Date) As String
    Dim lsCad As String
    Dim lsFechas As String
    
    lsCad = ""
    lsFechas = " entre " & Format(pdFecIni, gsFormatoFechaView) & " y " & Format(pdFecFin, gsFormatoFechaView) & " "
    
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro Aperturadas SOLES" & lsFechas & "- Con Orden de Pago -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro Aperturadas SOLES" & lsFechas & "- Sin Orden de Pago -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Plazo Fijo SOLES" & lsFechas, Moneda.gMonedaNacional, Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaApeAho("Cuentas de CTS SOLES" & lsFechas, Moneda.gMonedaNacional, Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro DOLARES" & lsFechas & "- Con Orden de Pago -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro DOLARES" & lsFechas & "- Sin Orden de Pago -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Plazo Fijo DOLARES" & lsFechas, Moneda.gMonedaExtranjera, Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaApeAho("Cuentas de CTS DOLARES" & lsFechas, Moneda.gMonedaExtranjera, Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    RepCtaApe = lsCad
End Function

Private Function GetCtaApeAho(psTit As String, psMoneda As Moneda, psProd As String, lnPagina As Long, lnItem As Long, pdFecIni As Date, pdFecFin As Date, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 28
    Dim lnMonto  As Currency
    Dim lnMontoApe  As Currency
    Dim lsFecha As String * 24
    Dim lsTipPers As String * 15
    Dim lsContador As String * 4
    Dim lsUsu As String * 4
    Dim lsMonto As String
    Dim lsMontoApe As String
    Dim lsCadena As String
    
    Dim lnContador As Integer
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    lsCadena = ""
    
    sqlAho = DevQry(psProd, pdFecIni, pdFecFin, psOrdPag, psMoneda, pbInactivas)
    Set rsAho = oCon.CargaRecordSet(sqlAho)
             
    If Not (rsAho.EOF And rsAho.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;20; ;9;Saldo.Disp;16; ;2;Fecha de Apertura;20;Personeria;16;Monto_Ape;18;Usu;6;", lnItem)
        
        lsCodCtaAnt = ""
        lnContador = 0
        
        While Not rsAho.EOF
            
            If rsAho!Cuenta <> lsCodCtaAnt Then
                lnContador = lnContador + 1
                lsContador = FillNum(Str(lnContador), 4, "0")
                lnMonto = rsAho!Monto
                lnMontoApe = rsAho!MontoApe
                lsMonto = Format$(lnMonto, "#,##0.00")
                lsMontoApe = Format$(lnMontoApe, "#,##0.00")
                lsUsu = rsAho!cUsu
            Else
                lsContador = ""
                lnMonto = 0
                lnMontoApe = 0
                lsMonto = ""
                lsMontoApe = ""
                lsUsu = ""
            End If
                
            lsCodCta = IIf(rsAho!Cuenta = lsCodCtaAnt, "", rsAho!Cuenta)
            lsNomPers = PstaNombre(rsAho!Nombre)
            lsFecha = Format(rsAho!FechaApe, gsFormatoFechaHoraView)
            lsTipPers = rsAho!Tipo
            
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(2) & Space(15 - Len(lsMonto)) & lsMonto & Space(2) & lsFecha _
                     & lsTipPers & Space(15 - Len(lsMontoApe)) & lsMontoApe & "  " & lsUsu & Chr(10)
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;20; ;5;Saldo.Disp;16; ;2;Fecha de Apertura;20; ;4;Personeria;20;Monto_Ape;18;Usu;4;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & Chr(12)
    End If
    
    GetCtaApeAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function

Private Function DevQry(psCodProd As String, dFecIni As Date, dFecFin As Date, Optional psOrdPag As String = "S", Optional psMoneda As Moneda = gMonedaNacional, Optional pbInactivas As Boolean = False) As String
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And PR.cCtaCod Like '___" & lsAgeCod & "%' "
    End If
    
    
    If pbInactivas Then
        DevQry = " Select Cuenta = A.cCodCta, Nombre = P.cNomPers, Monto = A.nSaldCntAC, FechaApe = A.dAperAC, " _
               & " UltCon = dUltCntAC FROM Ahorroc A Inner Join Perscuenta PC" _
               & " Inner Join Persona P On P.cCodPers = PC.cCodPers " _
               & " on A.cCodcTA = pC.cCodCta" _
               & " where bInactiva = 1 and cOrdPag = '" & psOrdPag & "' and substring(Aho.cCodCta,6,1) = '" & psMoneda & "' and cEstCtaAC not in ('C','U') order by Aho.cCodCta"
        Exit Function
    End If

    Select Case psCodProd
        Case gCapAhorros
            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
                   & " Case nPersoneria" _
                   & " When " & gPersonaNat & " then 'PERSONA NATURAL'" _
                   & " When " & gPersonaJurSFL & " then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFL & " then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFLCMAC & " then 'CMAC'" _
                   & " When " & gPersonaJurCFLCRAC & " then 'CRAC'" _
                   & " When " & gPersonaJurCFLCooperativa & " then 'COOPERATIVA' " _
                   & " When " & gPersonaJurCFLEdpyme & " then 'EDPYME' " _
                   & " End Tipo," _
                   & " IsNull(nMonto,0) MontoApe, ISNull((Select Right(cMovNro,4) From Mov M Where M.nMovnro = MC.nMovNro),'') cUsu" _
                   & " From Producto PR" _
                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                   & " Inner Join CaptacAhorros CAPAHO On PR.cCtaCod = CAPAHO.cCtaCod" _
                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & GetOpeGrupoRep(, 1, False) & "')" _
                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "' and CAPAHO.bOrdPag = " & psOrdPag & "" _
                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & "" _
                   & " Order By CAP.cCtaCod"
        Case gCapPlazoFijo
            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
                   & " Case nPersoneria" _
                   & " When " & gPersonaNat & " then 'PERSONA NATURAL'" _
                   & " When " & gPersonaJurSFL & " then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFL & " then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFLCMAC & " then 'CMAC'" _
                   & " When " & gPersonaJurCFLCRAC & " then 'CRAC'" _
                   & " When " & gPersonaJurCFLCooperativa & " then 'COOPERATIVA' " _
                   & " When " & gPersonaJurCFLEdpyme & " then 'EDPYME' " _
                   & " End Tipo," _
                   & " IsNull(nApertura,0) MontoApe, ISNull((Select Right(cMovNro,4) From Mov M Where M.nMovnro = MC.nMovNro),'') cUsu" _
                   & " From Producto PR" _
                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                   & " Inner Join CaptacPlazoFijo CAPPF On PR.cCtaCod = CAPPF.cCtaCod" _
                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & GetOpeGrupoRep(, 3, False) & "')" _
                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "'" _
                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & "" _
                   & " Order By CAP.cCtaCod"
        Case gCapCTS
            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
                   & " Case nPersoneria" _
                   & " When " & gPersonaNat & " then 'PERSONA NATURAL'" _
                   & " When " & gPersonaJurSFL & " then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFL & " then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                   & " When " & gPersonaJurCFLCMAC & " then 'CMAC'" _
                   & " When " & gPersonaJurCFLCRAC & " then 'CRAC'" _
                   & " When " & gPersonaJurCFLCooperativa & " then 'COOPERATIVA' " _
                   & " When " & gPersonaJurCFLEdpyme & " then 'EDPYME' " _
                   & " End Tipo," _
                   & " IsNull(nMonto,0) MontoApe, ISNull((Select Right(cMovNro,4) From Mov M Where M.nMovnro = MC.nMovNro),'') cUsu" _
                   & " From Producto PR" _
                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                   & " Inner Join CaptacCTS CAPCTS On PR.cCtaCod = CAPCTS.cCtaCod" _
                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & GetOpeGrupoRep(, 5, False) & "')" _
                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "'" _
                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" & sqlAdd & "" _
                   & " Order By CAP.cCtaCod"
    End Select
End Function

Public Function RepCtaCan(pdFecIni As Date, pdFecFin As Date) As String
    Dim lsCad As String
    Dim lsFechas As String
    
    lsCad = ""
    lsFechas = " entre " & Format(pdFecIni, gsFormatoFechaView) & " y " & Format(pdFecFin, gsFormatoFechaView) & " "
    
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro SOLES" & lsFechas & "- Con Orden de Pago -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro SOLES" & lsFechas & "- Sin Orden de Pago -", Moneda.gMonedaNacional, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Plazo Fijo SOLES" & lsFechas & "", Moneda.gMonedaNacional, Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de CTS SOLES" & lsFechas & "", Moneda.gMonedaNacional, Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro DOLARES" & lsFechas & "- Con Orden de Pago -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro DOLARES" & lsFechas & "- Sin Orden de Pago -", Moneda.gMonedaExtranjera, Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Plazo Fijo DOLARES" & lsFechas & "", Moneda.gMonedaExtranjera, Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de CTS DOLARES" & lsFechas & "", Moneda.gMonedaExtranjera, Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    RepCtaCan = lsCad
End Function

Private Function GetCtaCanAho(psTit As String, psMoneda As Moneda, psProd As String, lnPagina As Long, lnItem As Long, pdFecIni As Date, pdFecFin As Date, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As ADODB.Recordset
    Set rsAho = New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 28
    Dim lnMonto  As Currency
    Dim lsFecha As String * 24
    Dim lsTipPers As String * 30
    Dim lsContador As String * 4
    Dim lsCadena As String
    Dim lnContador As Integer
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsUsu As String
    
    oCon.AbreConexion
    
    Dim lsMonto As String
    lsCadena = ""
    
    sqlAho = DevQryCan(psProd, pdFecIni, pdFecFin, psOrdPag, psMoneda)
    Set rsAho = oCon.CargaRecordSet(sqlAho)
             
    If Not (rsAho.EOF And rsAho.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;25; ;5;Monto;16; ;2;Fecha de Cancelacion;22; ;4;Personeria;16;Usuario;18;", lnItem)
        
        lsCodCtaAnt = ""
        lnContador = 0
        
        While Not rsAho.EOF
            
            If rsAho!Cuenta <> lsCodCtaAnt Then
                lnContador = lnContador + 1
                lsContador = FillNum(Str(lnContador), 4, "0")
                lsCodCta = rsAho!Cuenta
                lnMonto = rsAho!Monto
                lsMonto = Space(18 - Len(Format(lnMonto, "#,##0.00"))) & Format(lnMonto, "#,##0.00")
                lsFecha = Format(rsAho!FechaCancel, gsFormatoFechaHoraView)
                lsUsu = rsAho!Usu
                lsTipPers = rsAho!Tipo
            Else
                lsContador = ""
                lsCodCta = ""
                lsMonto = Space(18)
                lsFecha = ""
                lsUsu = ""
                lsTipPers = ""
            End If
                
            
            lsNomPers = PstaNombre(rsAho!Nombre)
            
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(1) & lsMonto & Space(2) & lsFecha & lsTipPers & lsUsu & Chr(10)
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;25; ;5;Monto;16; ;2;Fecha de Apertura;20; ;4;Personeria;16;Usuario;4;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & Chr(12)
    End If
    
    GetCtaCanAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function

Private Function DevQryCan(psCodProd As String, dFecIni As Date, dFecFin As Date, Optional psOrdPag As String = "S", Optional psMoneda As Moneda = gMonedaNacional) As String
    Dim sqlAdd As String
    Dim sqlCodCan As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And P.cCtaCod Like '___" & lsAgeCod & "%' "
    End If
    ' MICA 22-Maro-2004
    Select Case psCodProd
        Case Producto.gCapAhorros
            sqlCodCan = "2004%"
        Case Producto.gCapPlazoFijo
            sqlCodCan = "2103%"
        Case Producto.gCapCTS
            sqlCodCan = "2204%"
    End Select
    
    ' Modificacion MICA 22/Marzo/2004
    
    Select Case psCodProd
        Case Producto.gCapAhorros
            
            DevQryCan = "Select Right(H.cMovNro,4) Usu, P.cCtaCod Cuenta, PE.cPersNombre Nombre," _
                        & " C.nSaldoDisp Monto, dbo.FechaHoraMov(H.cMovNro) FechaCancel," _
                        & " Case C.nPersoneria When 1 then 'PERSONA NATURAL' When 2 then 'PERSONA JURIDICA SIN FINES DE LUCRO' " _
                        & " When 3 then 'PERSONA JURIDICA CON FINES DE LUCRO' When 4 then 'CMAC' When 5 then 'CRAC' " _
                        & " When 6 then 'COOPERATIVA' When 7 then 'EDPYME' End  Tipo " _
                        & " From Producto P " _
                        & " Inner Join ProductoPersona PP On P.cCtaCod = PP.cCtaCod " _
                        & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod " _
                        & " Inner Join (select MC.cCtaCod, max(M.cMoVnro) cMovNro from MovCap MC inner join Mov m on M.nMovNro=MC.nMovNro " _
                        & " Where MC.cOpeCod like '" & sqlCodCan & "' group by MC.cCtaCod) H on H.cCtaCod=P.cCtaCod " _
                        & " Inner Join CaptacAhorros CA On P.cCtaCod = CA.cCtaCod " _
                        & " Inner Join Captaciones C On P.cCtaCod = C.cCtaCod " _
                        & " Where  Left(H.cMovNro,8) Between '" & Format(dFecIni, gsFormatoMovFecha) & "' And  '" & Format(dFecFin, gsFormatoMovFecha) & "'" _
                        & " And CA.bOrdPag = " & psOrdPag & " And substring(P.cCtaCod,9,1) = '" & psMoneda & "'" _
                        & " And Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & " Order By P.cCtaCod "

                      
        Case Producto.gCapPlazoFijo
            'DevQryCan = " Select  Right(CAP.cUltimaActualizacion,4) Usu, PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto," _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " dbo.FechaHoraMov(CAP.cUltimaActualizacion) FechaCancel," _
                      & " Case CAP.nPersoneria" _
                      & " When " & PersPersoneria.gPersonaNat & " then 'PERSONA NATURAL'" _
                      & " When " & PersPersoneria.gPersonaJurSFL & " then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                      & " When " & PersPersoneria.gPersonaJurCFL & " then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                      & " When " & PersPersoneria.gPersonaJurCFLCMAC & " then 'CMAC'" _
                      & " When " & PersPersoneria.gPersonaJurCFLCRAC & " then 'CRAC'" _
                      & " When " & PersPersoneria.gPersonaJurSFLFoncodes & " then 'FONCODES' " _
                      & " End  Tipo" _
                      & " From Producto PR" _
                      & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                      & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                      & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                      & " Inner Join CaptacPlazoFijo CAPPF On PR.cCtaCod = CAPPF.cCtaCod" _
                      & " Where  CAP.cUltimaActualizacion Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'" _
                      & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
                      & " And Left(PR.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & " Order By PR.cCtaCod"
        DevQryCan = " Select Right(H.cMovNro,4) Usu, P.cCtaCod Cuenta, PE.cPersNombre Nombre, " _
                    & " C.nSaldoDisp Monto, dbo.FechaHoraMov(H.cMovNro) FechaCancel, " _
                    & " Case C.nPersoneria When 1 then 'PERSONA NATURAL' When 2 then 'PERSONA JURIDICA SIN FINES DE LUCRO' " _
                    & " When 3 then 'PERSONA JURIDICA CON FINES DE LUCRO' When 4 then 'CMAC' When 5 then 'CRAC' " _
                    & " When 6 then 'COOPERATIVA' When 7 then 'EDPYME' End  Tipo From Producto P " _
                    & " Inner Join ProductoPersona PP On P.cCtaCod = PP.cCtaCod " _
                    & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod " _
                    & " Inner Join (select MC.cCtaCod, max(M.cMoVnro) cMovNro from MovCap MC inner join Mov m on M.nMovNro=MC.nMovNro " _
                    & " Where MC.cOpeCod like '" & sqlCodCan & "' group by MC.cCtaCod) H on H.cCtaCod=P.cCtaCod " _
                    & " Inner Join Captaciones C On P.cCtaCod = C.cCtaCod " _
                    & " Inner Join CaptacPlazoFijo CPF on CPF.cCtaCod= C.cCtaCod " _
                    & " Where Left(H.cMovNro,8) Between '" & Format(dFecIni, gsFormatoMovFecha) & "' And '" & Format(dFecFin, gsFormatoMovFecha) & "'" _
                    & " And substring(P.cCtaCod,9,1) = '" & psMoneda & "'" _
                    & " And Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & " Order By P.cCtaCod"

        Case Producto.gCapCTS
            'DevQryCan = " Select  Right(CAP.cUltimaActualizacion,4) Usu, PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto," _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " " _
                      & " dbo.FechaHoraMov(CAP.cUltimaActualizacion) FechaCancel," _
                      & " Case CAP.nPersoneria" _
                      & " When " & PersPersoneria.gPersonaNat & " then 'PERSONA NATURAL'" _
                      & " When " & PersPersoneria.gPersonaJurSFL & " then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                      & " When " & PersPersoneria.gPersonaJurCFL & " then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                      & " When " & PersPersoneria.gPersonaJurCFLCMAC & " then 'CMAC'" _
                      & " When " & PersPersoneria.gPersonaJurCFLCRAC & " then 'CRAC'" _
                      & " When " & PersPersoneria.gPersonaJurSFLFoncodes & " then 'FONCODES' " _
                      & " End  Tipo" _
                      & " From Producto PR" _
                      & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                      & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                      & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                      & " Inner Join CaptacCTS CAPCTS On PR.cCtaCod = CAPCTS.cCtaCod" _
                      & " Where  CAP.cUltimaActualizacion between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'" _
                      & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
                      & " And Left(PR.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & " Order By PR.cCtaCod"
       DevQryCan = " Select Right(H.cMovNro,4) Usu, P.cCtaCod Cuenta, PE.cPersNombre Nombre, " _
                    & " C.nSaldoDisp Monto, dbo.FechaHoraMov(H.cMovNro) FechaCancel, " _
                    & " Case C.nPersoneria When 1 then 'PERSONA NATURAL' When 2 then 'PERSONA JURIDICA SIN FINES DE LUCRO' " _
                    & " When 3 then 'PERSONA JURIDICA CON FINES DE LUCRO' When 4 then 'CMAC' When 5 then 'CRAC' " _
                    & " When 6 then 'COOPERATIVA' When 7 then 'EDPYME' End  Tipo From Producto P " _
                    & " Inner Join ProductoPersona PP On P.cCtaCod = PP.cCtaCod " _
                    & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod " _
                    & " Inner Join (select MC.cCtaCod, max(M.cMoVnro) cMovNro from MovCap MC inner join Mov m on M.nMovNro=MC.nMovNro " _
                    & " Where MC.cOpeCod like '" & sqlCodCan & "' group by MC.cCtaCod) H on H.cCtaCod=P.cCtaCod " _
                    & " Inner Join Captaciones C On P.cCtaCod = C.cCtaCod " _
                    & " Inner Join CaptacCTS CTS on CTS.cCtaCod= C.cCtaCod " _
                    & " Where Left(H.cMovNro,8) Between '" & Format(dFecIni, gsFormatoMovFecha) & "' And '" & Format(dFecFin, gsFormatoMovFecha) & "'" _
                    & " And substring(P.cCtaCod,9,1) = '" & psMoneda & "'" _
                    & " And Left(P.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " & sqlAdd & " Order By P.cCtaCod"
    End Select
    
End Function

Private Function CuentasInactivas(pMoneda As String, dFecIni As Date, dFecFin As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim TDesc As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    SQL = " SELECT ISNULL(CANT.cCtaCodAnt,'') cCtaCodAnt, CAPAHO.cCtaCod cCodCta, MC.nSaldoDisponible nSaldDispAC, CAPAHO.nSaldoAnterior nSaldAntAC, MC.nMonto nMonTran, CAPAHO.dUltContacto dUltCntAC" _
        & " FROM CaptacAhorros CAPAHO" _
        & " Inner Join MovCap MC On CAPAHO.cCtaCod = MC.cCtaCod" _
        & " Inner Join Mov M On MC.nMovNro = M.nMovNro" _
        & " Left  Join CaptacCtaAntRel CANT On  CAPAHO.cCtaCod = CANT.cCtaCod" _
        & " WHERE MC.cOpeCod in ('" & gAhoCancInact & "','" & gAhoEstActInac & "') And Substring(CAPAHO.cCtaCod,9,1) = '" & pMoneda & "'" _
        & " And M.nMovFlag <> " & gMovFlagExtornado & "" _
        & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " + ArmaFecha(gdFecSis) & Space(4) & "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
     lsCad = lsCad & Space(20) & "LISTADO DE CUENTAS INACTIVAS EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
     lsCad = lsCad & "CUENTA " & Space(5) & "CUENTAANT" & Space(5) & "S.ANTERIOR" & Space(5) & "S.DISPONIBLE" & Space(3) & "R.CAPITAL" & Space(2) & "F.ULTCONTACTO" & Space(5) + Chr(10)
     lsCad = lsCad & String(90, "=") + Chr(10)
     Do While Not rs.EOF
        CuentaA = rs!cCtaCodAnt
        RCapital = Abs(rs!nMonto)
        lsCad = lsCad & rs!cCodCta & Space(2) + CuentaA & JDNum(rs!nsaldantac, 13, True, 11, 2) & Space(2)
        lsCad = lsCad & JDNum(rs!nSaldDispAC, 13, True, 11, 2) & Space(4) & JDNum(Str(RCapital), 6, True, 4, 2) & Space(8) & Format(rs!dUltcntAC, gsFormatoFechaView) + Chr(10)
        TDesc = TDesc + RCapital
        rs.MoveNext
        vLineas = vLineas + 1
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(10) & "Resumen :" & Space(38) & JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " + ArmaFecha(gdFecSis) & Space(4) & "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
            lsCad = lsCad & Space(20) & "LISTADO DE CUENTAS INACTIVAS EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & "CUENTA " & Space(5) & "CUENTAANT" & Space(5) & "S.ANTERIOR" & Space(5) & "S.DISPONIBLE" & Space(3) & "R.CAPITAL" & Space(2) & "F.ULTCONTACTO" & Space(5) + Chr(10)
            lsCad = lsCad & String(90, "=") + Chr(10)
            vLineas = 1
        End If
     Loop
       lsCad = lsCad & Chr(10) & "Total Resumen :" & Space(32) & JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
    End If
    rs.Close
    Set rs = Nothing
    CuentasInactivas = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function


Private Function CuentasInmobilizadas(pMoneda As String, dFecIni As Date, dFecFin As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim TDesc As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    SQL = " Select MC.cCtaCod, PRD.nSaldo, 0 Interes, dbo.FechaMov(M.cMovNro) FFin, " _
        & " (Select dbo.FechaMov(Max(MA.cMovNro)) From Mov MA" _
        & " Inner Join MovCap MCA On MA.nMovNro = MCA.nMovNro" _
        & " Where MA.cOpeCod = '200803' And MA.nMovFlag = " & MovFlag.gMovFlagVigente & " And MCA.cCtaCod = MC.cCtaCod) FIni" _
        & " From Mov M" _
        & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
        & " Inner Join Producto PRD On PRD.cCtaCod = MC.cCtaCod" _
        & " Where M.cOpeCod = '200406'" _
        & " And Left(M.cMovNro,8) Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "' And M.nMovFlag = " & MovFlag.gMovFlagVigente & "" _
        & " And Substring(MC.cCtaCod,9,1) = '" & pMoneda & "'"

    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " + ArmaFecha(gdFecSis) & Space(4) & "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
     lsCad = lsCad & Space(20) & "LISTADO DE CUENTAS INACTIVAS EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
     lsCad = lsCad & "CUENTA " & Space(5) & "CUENTAANT" & Space(5) & "S.ANTERIOR" & Space(5) & "S.DISPONIBLE" & Space(3) & "R.CAPITAL" & Space(2) & "F.ULTCONTACTO" & Space(5) + Chr(10)
     lsCad = lsCad & String(90, "=") + Chr(10)
     Do While Not rs.EOF
        CuentaA = rs!cCtaCodAnt
        RCapital = Abs(rs!nMonto)
        lsCad = lsCad & rs!cCodCta & Space(2) + CuentaA & JDNum(rs!nsaldantac, 13, True, 11, 2) & Space(2)
        lsCad = lsCad & JDNum(rs!nSaldDispAC, 13, True, 11, 2) & Space(4) & JDNum(Str(RCapital), 6, True, 4, 2) & Space(8) & Format(rs!dUltcntAC, gsFormatoFechaView) + Chr(10)
        TDesc = TDesc + RCapital
        rs.MoveNext
        vLineas = vLineas + 1
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(10) & "Resumen :" & Space(38) & JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " + ArmaFecha(gdFecSis) & Space(4) & "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
            lsCad = lsCad & Space(20) & "LISTADO DE CUENTAS INACTIVAS EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & "CUENTA " & Space(5) & "CUENTAANT" & Space(5) & "S.ANTERIOR" & Space(5) & "S.DISPONIBLE" & Space(3) & "R.CAPITAL" & Space(2) & "F.ULTCONTACTO" & Space(5) + Chr(10)
            lsCad = lsCad & String(90, "=") + Chr(10)
            vLineas = 1
        End If
     Loop
       lsCad = lsCad & Chr(10) & "Total Resumen :" & Space(32) & JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
    End If
    rs.Close
    Set rs = Nothing
    CuentasInmobilizadas = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function



Public Function GetOpeGrupoRep(Optional pnGrupoCtaMov As CuentasMovidas = gCtasMovAhoApeCheque, Optional pnGrupoEstados As Integer = 0, Optional pbCuentasMovidas As Boolean = True) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsCadena As String
    
    If pbCuentasMovidas Then
        SQL = " Select cOpeCod From OpeTpoGrupo where nGrupoCtaMov = " & pnGrupoCtaMov
    Else
        SQL = " Select cOpeCod From OpeTpoGrupo where nGrupoApeCanCtas = " & pnGrupoEstados
    End If
    
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(SQL)
    lsCadena = ""
    
    If rs.EOF And rs.BOF Then
        GetOpeGrupoRep = ""
    Else
        While Not rs.EOF
            If lsCadena = "" Then
                lsCadena = rs.Fields(0)
            Else
                lsCadena = lsCadena & "','" & rs.Fields(0)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    rs.Close
    oCon.CierraConexion
    GetOpeGrupoRep = lsCadena
End Function

Public Function ImprimeRepGiros(ByVal dInicio As Date, ByVal dFin As Date, _
        ByVal sConvenio As String, ByVal sTitulo As String, Optional pbCancelacion As Boolean = False) As String
    Dim rsMov As ADODB.Recordset
    Dim nTotMonto As Double, nTotComision As Double
    Dim nCarLin As Integer
    Dim sTitRp1 As String
    Dim sTitRp2 As String
    Dim sMoneda As String, sNumPag As String, sFchtra As String
    Dim nLinPag As Integer, nCntPag As Integer
    Dim sFecha As String * 10
    Dim sCuenta As String * 12
    Dim sRemitente As String * 35, sDestinatario As String * 35
    Dim sAgencia As String * 15
    Dim nItem As Integer
    Dim sCad As String, sCuentaAux As String, sItem As String * 4
    Dim sMonto As String * 12, sComision As String * 12
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim i As Moneda
    Dim SQL As String
    Dim lsOpeCodGiro As String
    
    If Not pbCancelacion Then
        lsOpeCodGiro = "310[12]"
    Else
        lsOpeCodGiro = "310[3]"
    End If
    
    If Not oCon.AbreConexion Then
        ImprimeRepGiros = "No Existe una conexion activa con la Agencia Principal"
        Exit Function
    End If
    
    Set rsMov = New ADODB.Recordset
    rsMov.CursorLocation = adUseClient
    sCad = ""
    For i = Moneda.gMonedaNacional To Moneda.gMonedaExtranjera
        nTotMonto = 0
        nTotComision = 0
        
         SQL = " Select A.cCtaCod, A.dVigencia, A.cRemitente, A.nPrdEstado, A.cDestinatario, A.cAgencia, A.cCodDest, SUM(nMonto) nMonto, SUM(nComision) nComision" _
             & " From (Select G.cCtaCod, G.dVigencia, P.cPersNombre cRemitente, Prd.nPrdEstado, D.cDestinatario," _
             & "        LTRIM(RTRIM(TC.cAgeDescripcion)) cAgencia, D.cCodDest, nMonto = ISNULL(CASE WHEN M.cOpeCod LIKE '" & lsOpeCodGiro & "%'" _
             & "        THEN Abs(T.nMonto) END,0), nComision = ISNULL(CASE WHEN M.cOpeCod LIKE '" & lsOpeCodGiro & "%' THEN Abs(T.nMonto) END,0)" _
             & "     From Persona P" _
             & " INNER JOIN ProductoPersona PG" _
             & " Inner Join Producto Prd On PG.cCtaCod = Prd.cCtaCod" _
             & " INNER JOIN Giro G" _
             & " Inner Join" _
             & "    (Select G.cCtaCod, cDestinatario = Convert(Varchar(120),P.cPersNombre), P.cPersCod cCodDest" _
             & "      From Persona P INNER JOIN ProductoPersona PG" _
             & "            INNER JOIN Giro G ON PG.cCtaCod = G.cCtaCod ON P.cPersCod = PG.cPersCod" _
             & "        Where PG.nPrdPersRelac = 11" _
             & "        Union" _
             & "        Select G.cCtaCod, D.cDestinatario, cCodDest = '' From GiroDestinatario D" _
             & "            INNER JOIN Giro G ON D.cCtaCod = G.cCtaCod) D" _
             & "            INNER JOIN MovGiro T ON D.cCtaCod = T.cCtaCod ON G.cCtaCod = D.cCtaCod ON PG.cCtaCod = G.cCtaCod ON P.cPersCod = PG.cPersCod" _
             & "            INNER JOIN Mov M ON M.nMovNro = T.nMovNro" _
             & "            INNER JOIN Agencias TC ON G.cAgenciaDest = TC.cAgeCod" _
             & "            INNER JOIN Constante TC1 ON G.nPrdCtaTpo = RTRIM(TC1.nConsValor)" _
             & "        Where PG.nPrdPersRelac = 10 And TC1.nConsCod LIKE '1008'" _
             & "            And M.nMovFlag = 0 And M.cOpeCod" _
             & "        LIKE '" & lsOpeCodGiro & "%' And T.cCtaCod LIKE '________" & i & "%' And (M.cMovNro Between '" & Format(dInicio, gsFormatoMovFecha) & "' And '" & Format(dFin, gsFormatoMovFecha) & "')" _
             & "    ) A Group by A.cCtaCod, A.dVigencia, A.cRemitente, A.nPrdEstado, A.cDestinatario, A.cAgencia, A.cCodDest" _
             & " Order by A.cCtaCod"
        
        
        If rsMov.State = adStateOpen Then rsMov.Close
        
        Set rsMov = oCon.CargaRecordSet(SQL)
        Set rsMov.ActiveConnection = Nothing
        
        If Not (rsMov.EOF And rsMov.BOF) Then
            sNumPag = ""
            nLinPag = 0
            nCarLin = 145
            sTitRp1 = sTitulo
            sTitRp2 = "DEL " & Format$(dInicio, gsFormatoFechaView) & " AL " & Format$(dFin, gsFormatoFechaView)
            nCntPag = 1
            nItem = 0
            sMoneda = IIf(i = 1, "SOLES", "DOALRES")
            sNumPag = FillNum(Trim(Str(nCntPag)), 4, " ")
            If sCad <> "" Then sCad = sCad & Chr(12)
            sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, sNumPag) & Chr(10)
            sCad = sCad & String(nCarLin, "=") & Chr(10)
            sCad = sCad & "ITEM FECHA      CODIGO       REMITENTE                            DESTINATARIO                                MONTO     COMISION  AGE.DEST." & Chr(10)
            sCad = sCad & String(nCarLin, "=") & Chr(10)
            nLinPag = 7
            sCuentaAux = ""
            Do While Not rsMov.EOF
                If sCuentaAux <> rsMov("cCtaCod") Then
                    nItem = nItem + 1
                    RSet sItem = Trim(nItem)
                    sFecha = Format$(rsMov("dVigencia"), gsFormatoFechaView)
                    sCuenta = rsMov("cCtaCod")
                    sRemitente = PstaNombre(rsMov("cRemitente"), False)
                    RSet sMonto = Format$(rsMov("nMonto"), "#,##0.00")
                    RSet sComision = Format$(rsMov("nComision"), "#,##0.00")
                    sAgencia = rsMov("cAgencia")
                Else
                    sItem = ""
                    sFecha = ""
                    sCuenta = ""
                    sRemitente = ""
                    sMonto = ""
                    sComision = ""
                    sAgencia = ""
                End If
                sCuentaAux = rsMov("cCtaCod")
                sDestinatario = IIf(rsMov("cCodDest") = "", rsMov("cDestinatario"), PstaNombre(rsMov("cDestinatario"), False))
                If nLinPag > 58 Then
                    sCad = sCad & Chr(12)
                    nCntPag = nCntPag + 1
                    sNumPag = FillNum(Trim(Str(nCntPag)), 4, " ")
                    sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, sNumPag) & Chr(10)
                    sCad = sCad & String(nCarLin, "=") & Chr(10)
                    sCad = sCad & "ITEM FECHA      CODIGO       REMITENTE                            DESTINATARIO                                MONTO     COMISION  AGE.DEST." & Chr(10)
                    sCad = sCad & String(nCarLin, "=") & Chr(10)
                    nLinPag = 7
                End If
                sCad = sCad & sItem & " " & sFecha & " " & sCuenta & " " & sRemitente & "  " & sDestinatario & "  " & sMonto & " " & sComision & "  " & sAgencia & Chr(10)
                nLinPag = nLinPag + 1
                nTotMonto = nTotMonto + rsMov("nMonto")
                nTotComision = nTotComision + rsMov("nComision")
                rsMov.MoveNext
            Loop
            RSet sMonto = Format$(nTotMonto, "#,##0.00")
            RSet sComision = Format$(nTotComision, "#,##0.00")
            sCad = sCad & String(nCarLin, "-") & Chr(10)
            sCad = sCad & String(96, " ") & "TOTAL  " & sMonto & "  " & sComision & Chr(10)
            sCad = sCad & String(nCarLin, "-") & Chr(10)
        Else
            sCad = sCad & ""
        End If
    Next i
    rsMov.Close
    Set rsMov = Nothing
    ImprimeRepGiros = sCad
End Function

Private Function PlazosFijoCancelayMontos(pMoneda As String, dFecIni As Date, dFecFin As Date, pnMontoIni As Currency, pnMontoFin As Currency) As String
    Dim SQL As String, sCuentaAux As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim lnTotal As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsCtaCod As String * 18
    Dim lsSaldo As String * 15
    Dim lsApertura As String * 10
    Dim lsVencimiento As String * 10
    Dim lsPlazo As String * 5
    Dim lsNombre As String * 25
    Dim lsDireccion As String * 25
    Dim lsFono As String * 6
    Dim sqlAdd As String
    Dim lsCtaCodAnt As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Left(CAP.cCtaCod,5) Like '___" & lsAgeCod & "'"
    End If
    
    oCon.AbreConexion
    
    SQL = " Select PRD.cCtaCod, PRD.nSaldo, CAP.dApertura, PF.nPlazo, DateAdd(day,PF.nPlazo,CAP.dApertura) FFin, " _
        & "     VW.cPersNombre , VW.cPersDireccDomicilio, ISNULL(VW.cPersTelefono,'') cPersTelefono" _
        & " From Producto PRD" _
        & " Inner Join Captaciones CAP ON PRD.cCtaCod = CAP.cCtaCod" _
        & " Inner Join CaptacPlazoFijo PF ON CAP.cCtaCod = PF.cCtaCod" _
        & " Inner Join (" _
        & "         Select PE.cPersNombre, PE.cPersDireccDomicilio, cPersTelefono, PP.cCtaCod From ProductoPersona PP" _
        & "             Inner Join Persona PE On PP.cPersCod = PE.cPersCod" _
        & "             Where nPrdPersRelac = 10) As VW On VW.cCtaCod = PF.cCtaCod" _
        & " Where DateAdd(day,PF.nPlazo,CAP.dApertura) Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "' And nSaldo Between " & pnMontoIni & " And " & pnMontoFin & "" _
        & " " & sqlAdd & " And Left(PRD.nPrdEstado,2) Not In (" & Left(CaptacEstado.gCapEstAnulada, 2) & "," & Left(CaptacEstado.gCapEstCancelada, 2) & ") And Substring(CAP.cCtaCod,9,1) = '" & pMoneda & "' "
    SQL = SQL & " ORDER BY PRD.cCtaCod "
    Set rs = oCon.CargaRecordSet(SQL)
    'dFecIni , dFecFin, pnMontoIni, pnMontoFin


    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & lsEmpresa & " " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
     lsCad = lsCad & Space(10) & "LISTADO DE PLAZO FIJO (ANALISIS) EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") & " Entre :  " & Format(dFecIni, gsFormatoFecha) & " - " & Format(dFecFin, gsFormatoFecha) & " y " & Format(pnMontoIni, "#,##0.00") & " - " & Format(pnMontoFin, "#,##0.00") & Chr(10) & Chr(10)
     lsCad = lsCad & String(130, "=") & Chr(10)
     lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(15) & "DIRECCION" & Space(12) & "TLFNO" & Space(12) & "MONTO" & Space(5) & "F.INI." & Space(5) & "F.FIN." & Space(5) & "PLAZO" & Chr(10)
     lsCad = lsCad & String(130, "=") & Chr(10)
     sCuentaAux = ""
     Do While Not rs.EOF
        
        If lsCtaCodAnt <> rs!Cctacod Then
            lsCtaCod = rs!Cctacod
            RSet lsSaldo = Format(rs!nSaldo, "#,##0.00")
            lsApertura = Format(rs!dApertura, gsFormatoFechaView)
            lsVencimiento = Format(rs!dApertura, gsFormatoFechaView)
            RSet lsPlazo = rs!nPlazo
            lsNombre = rs!cPersNombre
            lsDireccion = rs!cPersDireccDomicilio
            lsFono = rs!cPersTelefono
            lnTotal = lnTotal + rs!nSaldo
        Else
            lsCtaCod = ""
            RSet lsSaldo = ""
            lsApertura = ""
            lsVencimiento = ""
            RSet lsPlazo = ""
            lsNombre = rs!cPersNombre
            lsDireccion = rs!cPersDireccDomicilio
            lsFono = rs!cPersTelefono
        End If
                
                
        lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsDireccion & Space(2) & lsFono & Space(2) & lsSaldo & Space(2) & lsApertura & Space(2) & lsVencimiento & Space(2) & lsPlazo & Chr(10)
        
        
        lsCtaCodAnt = rs!Cctacod
        
        rs.MoveNext
        vLineas = vLineas + 1
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(10) & "Resumen :" & Space(38) & JDNum(Str(lnTotal), 13, True, 11, 2) & Chr(10)
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & lsEmpresa & " " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
            lsCad = lsCad & Space(20) & "LISTADO DE PLAZO FIJO (ANALISIS) EN " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & String(130, "=") & Chr(10)
            lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(15) & "DIRECCION" & Space(12) & "TLFNO" & Space(12) & "MONTO" & Space(5) & "F.INI." & Space(5) & "F.FIN." & Space(5) & "PLAZO" & Chr(10)
            lsCad = lsCad & String(130, "=") & Chr(10)
            vLineas = 1
        End If
     Loop
    lsCtaCod = ""
    RSet lsSaldo = Format(lnTotal, "#,##0.00")
    lsApertura = ""
    lsVencimiento = ""
    lsPlazo = ""
    lsNombre = ""
    lsDireccion = ""
    lsFono = ""
    lsCad = lsCad & String(130, "=") & Chr(10)
    lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsDireccion & Space(2) & lsFono & Space(2) & lsSaldo & Space(2) & lsApertura & Space(2) & lsVencimiento & Space(2) & lsPlazo & Chr(10)

    End If
    rs.Close
    Set rs = Nothing
    PlazosFijoCancelayMontos = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Private Function PlazosFijoCancelayMontosEspecial(pMoneda As String, dFecIni As Date, dFecFin As Date, pnMontoIni As Currency, pnMontoFin As Currency) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim lnTotal As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsCtaCod As String * 18
    Dim lsSaldo As String * 15
    Dim lsApertura As String * 10
    Dim lsVencimiento As String * 10
    Dim lsPlazo As String * 5
    Dim lsNombre As String * 40
    Dim lsDireccion As String * 40
    Dim lsFono As String * 6
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Left(CAP.cCtaCod,5) Like '___" & lsAgeCod & "'"
    End If
    
    
    oCon.AbreConexion
    
    SQL = " Select  PRD.cCtaCod, PRD.nSaldo, CAP.dApertura, PF.nPlazo, DateAdd(day,PF.nPlazo,CAP.dApertura) FFin, " _
        & "     VW.cPersNombre , VW.cPersDireccDomicilio, VW.cPersTelefono" _
        & " From Producto PRD" _
        & " Inner Join Captaciones CAP ON PRD.cCtaCod = CAP.cCtaCod" _
        & " Inner Join CaptacPlazoFijo PF ON CAP.cCtaCod = PF.cCtaCod" _
        & " Inner Join (" _
        & "         Select PE.cPersNombre, PE.cPersDireccDomicilio + '-' + (Select cUbiGeoDescripcion From UbicacionGeografica UG Where UG.cUbiGeoCod = PE.cPersDireccUbiGeo) cPersDireccDomicilio, cPersTelefono, PP.cCtaCod From ProductoPersona PP" _
        & "             Inner Join Persona PE On PP.cPersCod = PE.cPersCod" _
        & "             Where nPrdPersRelac = 10) As VW On VW.cCtaCod = PF.cCtaCod" _
        & " Where DateAdd(day,PF.nPlazo,CAP.dApertura) Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "' And nSaldo Between " & pnMontoIni & " And " & pnMontoFin & "" _
        & " " & sqlAdd & " And Left(PRD.nPrdEstado,2) Not In (" & Left(CaptacEstado.gCapEstAnulada, 2) & "," & Left(CaptacEstado.gCapEstCancelada, 2) & ") And Substring(CAP.cCtaCod,9,1) = '" & pMoneda & "' And nPrdTasaInteres = " & gCapTasaPreferencial
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
     lsCad = lsCad & Space(20) & "LISTADO DE PLAZO FIJO (TASA PREFERENCIAL) EN " & IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
     lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & "DIRECCION" & Space(40) & "MONTO" & Space(5) & "F.INI." & Space(5) & "F.FIN." & Space(5) & "PLAZO" & Chr(10)
     lsCad = lsCad & String(160, "=") & Chr(10)
     Do While Not rs.EOF
        lsCtaCod = rs!Cctacod
        RSet lsSaldo = Format(rs!nSaldo, "#,##0.00")
        lsApertura = Format(rs!dApertura, gsFormatoFechaView)
        lsVencimiento = Format(rs!dApertura, gsFormatoFechaView)
        RSet lsPlazo = rs!nPlazo
        lsNombre = rs!cPersNombre
        lsDireccion = rs!cPersDireccDomicilio
        lsFono = rs!cPersTelefono
        
        lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsDireccion & Space(2) & lsFono & Space(2) & lsSaldo & Space(2) & lsApertura & Space(2) & lsVencimiento & Space(2) & lsPlazo & Chr(10)
        
        lnTotal = lnTotal + rs!nSaldo
        
        rs.MoveNext
        vLineas = vLineas + 1
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(10) & "Resumen :" & Space(38) & JDNum(Str(lnTotal), 13, True, 11, 2) & Chr(10)
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
            lsCad = lsCad & Space(20) & "LISTADO DE PLAZO FIJO (TASA PREFERENCIAL) " + IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & "DIRECCION" & Space(40) & "MONTO" & Space(5) & "F.INI." & Space(5) & "F.FIN." & Space(5) & "PLAZO" & Chr(10)
            lsCad = lsCad & String(160, "=") & Chr(10)
            vLineas = 1
        End If
     Loop
    lsCtaCod = ""
    RSet lsSaldo = Format(lnTotal, "#,##0.00")
    lsApertura = ""
    lsVencimiento = ""
    lsPlazo = ""
    lsNombre = ""
    lsDireccion = ""
    lsFono = ""
    lsCad = lsCad & String(160, "=") & Chr(10)
    lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsDireccion & Space(2) & lsFono & Space(2) & lsSaldo & Space(2) & lsApertura & Space(2) & lsVencimiento & Space(2) & lsPlazo & Chr(10)
    
    End If
    rs.Close
    Set rs = Nothing
    PlazosFijoCancelayMontosEspecial = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Public Function AvisoRenovacionPF(ByVal pdFecha As Date, ByVal psTextoCarta As String, Optional ByVal psCodAge As String) As String
Dim SQL As String
Dim rs As Recordset
Dim oCon As DConecta
Dim lsCad As String
Dim lsCarta As String
Dim lsFechaIni As String
Dim lsFechaFin As String

lsFechaFin = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(pdFecha), 1, 2)
lsFechaIni = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(DateAdd("d", -1, pdFecha)), 1, 2)

Set oCon = New DConecta

    oCon.AbreConexion

    SQL = "SELECT C.cCtaCod, cPersNombre, cPersDireccDomicilio, cUbiGeoDescripcion, nPlazo, dbo.ConvierteTNAaTEA(nTasaInteres) nTasaInteres, nSaldo CapitalAct,  " _
        & "nSaldoDisponible - nMonto CapitalIni, nMonto Interes, DATEADD(dd, 0, dRenovacion) FechaVencimiento, " _
        & "DATEADD(dd, nPlazo, dRenovacion) FechaVencAct, " _
        & "nInteres = (SELECT SUM(nMonto) " _
        & "             FROM MOVCAP MC1 INNER JOIN MOV M1 ON MC1.nMovNro = M1.nMovNro " _
        & "                     INNER JOIN CAPTACPLAZOFIJO PF ON MC1.cCtacod = PF.cCtaCod " _
        & "                     WHERE M1.cOpeCod = 210201 AND nMovFlag <> 2 AND MC1.cCtaCod = C.cCtaCod " _
        & "                             AND SUBSTRING(M1.cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "') " _
        & "FROM PRODUCTO P INNER JOIN CAPTACIONES C ON P.cCtaCod = C.cCtaCod " _
        & "                INNER JOIN CAPTACPLAZOFIJO PF ON C.cCtaCod = PF.cCtaCod " _
        & "                INNER JOIN MOVCAP MC ON MC.cCtaCod = C.cCtaCod INNER JOIN MOV M ON MC.nMovNro = M.nMovNro " _
        & "                INNER JOIN PRODUCTOPERSONA PP ON PF.cCtaCod = PP.cCtaCod INNER JOIN PERSONA T ON PP.cPersCod = T.cPersCod " _
        & "                INNER JOIN UBICACIONGEOGRAFICA U ON T.cPersDireccUbiGeo = U.cUbiGeoCod " _
        & "WHERE dApertura <> dRenovacion AND DATEDIFF(dd, dRenovacion, '" & Format(pdFecha, gsFormatoFecha) & "') BETWEEN 0 AND 1 " _
        & "      AND M.cOpeCod = '210501' AND nMovFlag <> 2 AND nPrdPersRelac = 10 AND " _
        & " SUBSTRING(M.cMovNro, 1, 8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "'"
        
        If psCodAge <> "" Then
            SQL = SQL & " AND SUBSTRING(C.cCtaCod,4,2) = '" & psCodAge & "'"
        End If
        
        SQL = SQL & " ORDER BY C.cCtaCod "
        
    Set rs = oCon.CargaRecordSet(SQL)
            
    If Not (rs.EOF And rs.BOF) Then
            
        Do While Not rs.EOF
        
             lsCarta = psTextoCarta
             lsCarta = Replace(lsCarta, "<<FECHA>>", Format(pdFecha, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<NOMBRE>>", PstaNombre(rs!cPersNombre, False), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CUENTA>>", rs!Cctacod, , 2, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CAPITALINI>>", ImpreFormat(rs!CapitalIni, 7, 2, True), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<PLAZO>>", "      " + ImpreFormat(rs!nPlazo, 4, False), , 2, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<FECHAVENCANT>>", Format(rs!FechaVencimiento, gcFormatoFechaView), , 2, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<INTGANADOS>>", ImpreFormat(rs!Interes + IIf(IsNull(rs!nInteres), 0, rs!nInteres), 7, 2, True), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<CAPITALACT>>", ImpreFormat(rs!CapitalAct, 7, 2, True), , 2, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<FECHAVENCNUE>>", Format(rs!FechaVencAct, gcFormatoFechaView), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<TASAINTACT>>", ImpreFormat(rs!nTasaInteres, 3, 2, True), , 1, vbTextCompare)
             lsCarta = Replace(lsCarta, "<<DIRECCION>>", rs!cPersDireccDomicilio & " - " + rs!cUbiGeoDescripcion, , 1, vbTextCompare)
             
             lsCad = lsCad & lsCarta & Chr(12)
         
             rs.MoveNext
         Loop
                      
       AvisoRenovacionPF = lsCad
               
    End If
        
End Function
  
  
Public Function ProtocoloOperaciones(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, _
                            pgsEmpresa As String, pgdFecSis As Date, Optional pnMoneda As Moneda = gMonedaNacional, _
                            Optional psUsuario As String = "XXX", Optional pbSalto As Boolean = True, _
                            Optional psFecha As String = "", Optional psAgencia As String = "", Optional lsOrden As String = "", Optional lsCheck As String = "") As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim lsDoc As String * 18
    Dim lsCodCta As String
    Dim lsNomOpe As String * 33
    Dim lsCodUsu As String, sTitulo1 As String
    Dim lbBan As Boolean
    Dim lsCadena As String
    Dim lbHoy As Boolean
    Dim i As Long
    Dim lsFec As String
    Dim sqlAdd As String
    Dim ldIni As Date
    Dim ldFin As Date
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnLineas As Integer
    
    Dim nMonto As Double
    Dim nCantidad As Long
    Dim nCantidadT As Long
    Dim nMontoT As Double
    Dim sTEmpoOpe As String
    Dim sOperacionTempo As String * 33
    
    lbHoy = IIf(psFecha = Format$(pgdFecSis, gsFormatoFechaView) Or psFecha = "", True, False)
    
    If psAgencia <> "" Then
        sqlAdd = " And SubString(M.cMovNro,18,2) = '" & psAgencia & "'"
    Else
        sqlAdd = ""
    End If
    
    lbBan = False
    lsFec = Format$(CDate(psFecha), "dd mmm yyyy")
    
    ldIni = CDate(psFecha)
    ldFin = DateAdd("d", 1, CDate(psFecha))
    lnRango = 58
    
    oCon.AbreConexion
    
    'Ahorros
    If psUsuario = "XXX" Then
        sTitulo1 = psTitulo & " DEL " & UCase(lsFec) & " - "
    
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovColDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union"
            
            SQL = SQL & "   Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cNumRecibo cCtaCod, '' cDocNro" _
                & "         From MovServicios MC" _
                & "         Inner Join Mov M On MC.nMovNro = M.nMovNro" _
                & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
                & "     Group By M.nMovNro, M.cOpeCod, MC.cNumRecibo  ) As Ope On Ope.nMovNro = M.nMovNro" _
                & " Where (SubString(M.cOpeCod,1,1) = '2' Or M.cOpeCod like '3%') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "' " _
                & " And M.nMovFlag = 0 " & sqlAdd
                
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
                
                ' Where ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
    Else
        sTitulo1 = psTitulo & " - " & psUsuario & " - " & " DEL " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovColDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union"
            
            SQL = SQL & "   Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cNumRecibo cCtaCod, '' cDocNro" _
                & "         From MovServicios MC" _
                & "         Inner Join Mov M On MC.nMovNro = M.nMovNro" _
                & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
                & "     Group By M.nMovNro, M.cOpeCod, MC.cNumRecibo  ) As Ope On Ope.nMovNro = M.nMovNro" _
                & " Where (SubString(M.cOpeCod,1,1) = '2') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "'" _
                & " And M.nMovFlag = 0 " & sqlAdd & "  And  Right(M.cMovNro,4) = '" & psUsuario & "' "
                
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
                
            ' And ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
            
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not RSVacio(rs) Then
    
        PonerSalto lsCadena
        lbBan = True
        
        lsCadena = lsCadena & CabeceraPagina(sTitulo1 & " AHORRO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
        lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
        
        sTEmpoOpe = rs!copecod
        sOperacionTempo = rs!cNomOpe
        nMonto = 0
        nCantidad = 0
        
        While Not rs.EOF
            
            If lsOrden = "2" Then 'por codigo de operacion
                If sTEmpoOpe <> rs!copecod Then
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & String(131, "-") & Chr(10)
                    End If
                    
                    lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
                    
                    'lsCadena = lsCadena & ImpreFormat(nMonto, 10, 2)
                    
                    'lsCadena = lsCadena & Space(14 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(27) & " [" & sTEmpoOpe & "]" & Chr(10)
                    
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & Chr(10)
                    End If
                    
                    lsNomOpe = rs!cNomOpe
                    sOperacionTempo = rs!cNomOpe
                    nMonto = rs!nMonTran
                    sTEmpoOpe = rs!copecod
                    nCantidad = 1
                Else
                    nMonto = nMonto + rs!nMonTran
                    nCantidad = nCantidad + 1
                End If
                nMontoT = nMontoT + rs!nMonTran
                nCantidadT = nCantidadT + 1
            End If
            
            
            If IsNull(rs!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = rs!cNumDoc
            End If
            
            If IsNull(rs!cCodCta) Then
                lsCodCta = " "
            Else
                lsCodCta = rs!cCodCta
            End If
            
            If IsNull(rs!cCodUsu) Then
                If IsNull(rs!cCodusurem) Then
                    lsCodUsu = ""
                Else
                    lsCodUsu = rs!cCodusurem
                End If
            Else
                lsCodUsu = rs!cCodUsu
            End If

            i = 1

            lsNomOpe = rs!cNomOpe
            If lsCheck <> "1" Then
            lsCadena = lsCadena & Space(10 - Len(Str(rs!nNumTran))) & Str(rs!nNumTran) _
                                & Space(20 - Len(lsCodCta)) & lsCodCta _
                                & Space(2) & Trim(lsNomOpe) _
                                & Space(26 - Len(Trim(lsDoc)) + 20 - Len(Trim(lsNomOpe))) & Trim(lsDoc) _
                                & Space(16 - Len(Format(rs!nMonTran * i, "###,##0.00"))) & Format(rs!nMonTran * i, "###,##0.00") _
                                & Space(7 - Len(lsCodUsu)) & lsCodUsu _
                                & Space(7 - Len(rs!CCodAge)) & rs!CCodAge _
                                & Space(14 - Len(Format(rs!dFecTran, "hh:mm:ss AMPM"))) & Format(rs!dFecTran, "hh:mm:ss AMPM") & " [" & rs!copecod & "]" & Chr$(10)
            End If
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                lsCadena = lsCadena & Chr$(12)
                lsCadena = lsCadena + CabeceraPagina(sTitulo1 & " AHORRO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
                lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
            End If
            
            rs.MoveNext
        Wend
        
        If lsOrden = "2" Then 'por codigo de operacion
            If lsCheck <> "1" Then
                lsCadena = lsCadena & String(131, "-") & Chr(10)
            End If
            lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
        End If
        
        'If lsOrden = "2" Then 'por codigo de operacion
        '    lsCadena = lsCadena & String(122, "=") & Chr(10)
        '    sOperacionTempo = ""
        '    lsCadena = lsCadena & Space(26) & "TOTAL " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidadT, 6, 0) & ")" & Space(1) & ImpreFormat(nMontoT, 10, 2) & Chr(10) & Chr(10)
        'End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    nCantidadT = 0
    nMontoT = 0
    
    'Prendario
    If psUsuario = "XXX" Then
        sTitulo1 = psTitulo & " DEL " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCol MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where (SubString(M.cOpeCod,1,2) In ('12') Or M.cOpeCod like '32%') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "' " _
            & " And M.nMovFlag = 0 " & sqlAdd & "  " 'Order by nNumTran"
            
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
            
            'Where ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
    Else
        sTitulo1 = psTitulo & " - " & psUsuario & " - " & " del " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCol MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where (SubString(M.cOpeCod,1,2) In ('12') Or M.cOpeCod like '31%') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "'" _
            & " And M.nMovFlag = 0 " & sqlAdd & "  And  Right(M.cMovNro,4) = '" & psUsuario & "' " 'Order by nNumTran"
            
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
            
            'And ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
     End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not RSVacio(rs) Then
        PonerSalto lsCadena
        
        lbBan = True
        
        lsCadena = lsCadena + CabeceraPagina(sTitulo1 & " PRENDARIO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
        lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
        
        sTEmpoOpe = rs!copecod
        sOperacionTempo = rs!cNomOpe
        nMonto = 0
        nCantidad = 0
        
        While Not rs.EOF
            
            If lsOrden = "2" Then 'por codigo de operacion
                If sTEmpoOpe <> rs!copecod Then
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & String(131, "-") & Chr(10)
                    End If
                    lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & Chr(10)
                    End If
                    lsNomOpe = rs!cNomOpe
                    sOperacionTempo = rs!cNomOpe
                    nMonto = rs!nMonTran
                    sTEmpoOpe = rs!copecod
                    nCantidad = 1
                Else
                    nMonto = nMonto + rs!nMonTran
                    nCantidad = nCantidad + 1
                End If
                nMontoT = nMontoT + rs!nMonTran
                nCantidadT = nCantidadT + 1
            End If
            
            If IsNull(rs!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = rs!cNumDoc
            End If
            
            If IsNull(rs!cCodCta) Then
                lsCodCta = " "
            Else
                lsCodCta = rs!cCodCta
            End If
            
            If IsNull(rs!cCodUsu) Then
                If IsNull(rs!cCodusurem) Then
                    lsCodUsu = ""
                Else
                    lsCodUsu = rs!cCodusurem
                End If
            Else
                lsCodUsu = rs!cCodUsu
            End If
            
            lsNomOpe = rs!cNomOpe
            If lsCheck <> "1" Then
            lsCadena = lsCadena & Space(10 - Len(Str(rs!nNumTran))) & Str(rs!nNumTran) _
                                & Space(20 - Len(lsCodCta)) & lsCodCta _
                                & Space(2) & Trim(lsNomOpe) _
                                & Space(26 - Len(Trim(lsDoc)) + 20 - Len(Trim(lsNomOpe))) & Trim(lsDoc) _
                                & Space(16 - Len(Format(rs!nMonTran, "###,##0.00"))) & Format(rs!nMonTran, "###,##0.00") _
                                & Space(7 - Len(lsCodUsu)) & lsCodUsu _
                                & Space(7 - Len(rs!CCodAge)) & rs!CCodAge _
                                & Space(14 - Len(Format(rs!dFecTran, "hh:mm:ss AMPM"))) & Format(rs!dFecTran, "hh:mm:ss AMPM") & " [" & rs!copecod & "]" & Chr$(10)
            End If
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                lsCadena = lsCadena & Chr$(12)
                lsCadena = lsCadena & CabeceraPagina(sTitulo1 & " PRENDARIO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
                lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
            End If
            
            rs.MoveNext
        Wend
        
        If lsOrden = "2" Then 'por codigo de operacion
            lsCadena = lsCadena & String(131, "-") & Chr(10)
            lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
        End If
        
        'If lsOrden = "2" Then 'por codigo de operacion
        '    lsCadena = lsCadena & String(122, "=") & Chr(10)
        '    sOperacionTempo = ""
        '    lsCadena = lsCadena & Space(26) & "TOTAL " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidadT, 6, 0) & ")" & Space(1) & ImpreFormat(nMontoT, 10, 2) & Chr(10) & Chr(10)
        'End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    nCantidadT = 0
    nMontoT = 0
    
    'Creditos
    If psUsuario = "XXX" Then
        sTitulo1 = psTitulo & " DEL " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovColDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where (Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "')" _
            & " AND nPrdConceptoCod NOT IN (8102, 8503, 8504, 8502,8505)" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where (SubString(M.cOpeCod,1,2) In ('10','13','14','15') Or M.cOpeCod like '33%') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "' " _
            & " And M.nMovFlag = 0 " & sqlAdd & "  " 'Order by nNumTran"
        
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
        
        'Where ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
        
    Else
        sTitulo1 = psTitulo & " - " & psUsuario & " - " & " DEL " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovCapDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where nConceptoCod In (1,2) And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "     Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _
            & "         From MovColDet MC" _
            & "         Inner Join Mov M On MC.nMovNro = M.nMovNro And MC.cOpeCod = M.cOpeCod" _
            & "         Left Join  MovDoc MD On MD.nMovNro = MC.nMovNro" _
            & "         Where (Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "')" _
            & " AND nPrdConceptoCod NOT IN (8102, 8503, 8504, 8502, 8505) " _
            & "     Group By M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where (SubString(M.cOpeCod,1,2) In ('10','13','14','15') Or M.cOpeCod like '33%') And Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "'" _
            & " And M.nMovFlag = 0 " & sqlAdd & "  And  Right(M.cMovNro,4) = '" & psUsuario & "' " 'Order by nNumTran"
            
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
            
            'And ... and substring(M.cMovNro,18,2) = '" & psAgencia & "'
    End If
'            & "     Select Sum(Case nPrdConceptoCod WHEN 8502 THEN 0 WHEN 8505 THEN 0 WHEN 8508 THEN 0 ELSE nMonto END) Monto, M.nMovNro, M.cOpeCod, MC.cCtaCod, MD.cDocNro" _

    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not RSVacio(rs) Then
        PonerSalto lsCadena
        lbBan = True
        
        lsCadena = lsCadena + CabeceraPagina(sTitulo1 & " CREDITO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
        lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
        
        sTEmpoOpe = rs!copecod
        sOperacionTempo = rs!cNomOpe
        nMonto = 0
        nCantidad = 0
        
        While Not rs.EOF
            
            If lsOrden = "2" Then 'por codigo de operacion
                If sTEmpoOpe <> rs!copecod Then
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & String(131, "-") & Chr(10)
                    End If
                    lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & Chr(10)
                    End If
                    lsNomOpe = rs!cNomOpe
                    sOperacionTempo = rs!cNomOpe
                    nMonto = rs!nMonTran
                    sTEmpoOpe = rs!copecod
                    nCantidad = 1
                Else
                    nMonto = nMonto + rs!nMonTran
                    nCantidad = nCantidad + 1
                End If
                nMontoT = nMontoT + rs!nMonTran
                nCantidadT = nCantidadT + 1
            End If
            
            If IsNull(rs!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = rs!cNumDoc
            End If
            
            If IsNull(rs!cCodCta) Then
                lsCodCta = " "
            Else
                lsCodCta = rs!cCodCta
            End If
            
            If IsNull(rs!cCodUsu) Then
                If IsNull(rs!cCodusurem) Then
                    lsCodUsu = ""
                Else
                    lsCodUsu = rs!cCodusurem
                End If
            Else
                lsCodUsu = rs!cCodUsu
            End If
            
            lsNomOpe = Mid(rs!cNomOpe, 1, 26)
            If lsCheck <> "1" Then
            lsCadena = lsCadena & Space(10 - Len(Str(rs!nNumTran))) & Str(rs!nNumTran) _
                                & Space(20 - Len(lsCodCta)) & lsCodCta _
                                & Space(2) & Trim(lsNomOpe) _
                                & Space(26 - Len(Trim(lsDoc)) + 20 - Len(Trim(lsNomOpe))) & Trim(lsDoc) _
                                & Space(16 - Len(Format(rs!nMonTran, "###,##0.00"))) & Format(rs!nMonTran, "###,##0.00") _
                                & Space(7 - Len(lsCodUsu)) & lsCodUsu _
                                & Space(7 - Len(rs!CCodAge)) & rs!CCodAge _
                                & Space(14 - Len(Format(rs!dFecTran, "hh:mm:ss AMPM"))) & Format(rs!dFecTran, "hh:mm:ss AMPM") & " [" & rs!copecod & "]" & Chr$(10)
            End If
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                lsCadena = lsCadena & Chr$(12)
                lsCadena = lsCadena & CabeceraPagina(sTitulo1 & " CREDITO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
                lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
            End If
            
            rs.MoveNext
        Wend
        
        If lsOrden = "2" Then 'por codigo de operacion
            If lsCheck <> "1" Then
                lsCadena = lsCadena & String(131, "-") & Chr(10)
            End If
            lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
        End If
        
        'If lsOrden = "2" Then 'por codigo de operacion
        '    lsCadena = lsCadena & String(122, "=") & Chr(10)
        '    sOperacionTempo = ""
        '    lsCadena = lsCadena & Space(26) & "TOTAL " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidadT, 6, 0) & ")" & Space(1) & ImpreFormat(nMontoT, 10, 2) & Chr(10) & Chr(10)
        'End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    nCantidadT = 0
    nMontoT = 0
    
    'Otros
    If psUsuario = "XXX" Then
        sTitulo1 = psTitulo & " DEL " & UCase(lsFec) & " - "
        
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '000000002' cCtaCod, '' cDocNro     " _
            & "         From MovCompraVenta MC" _
            & "             Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & " Group By M.nMovNro, M.cOpeCod" _
            & " Union" _
            & " Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))) cCtaCod, Mc.cNroDoc" _
            & "     From MovOpeVarias MC " _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro " _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cNroDoc" _
            & " Union" _
            & "     Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))) cCtaCod, 'Dest : ' + MC.cUsuDest cNroDoc" _
            & "     From MovHabDevCajero MC" _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cUsuDest"
            
        SQL = SQL & " Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, '00000000' + LTrim(RTrim(Str(MC.nMoneda))) cCtaCod, MC.cNumRecibo cNroDoc" _
            & "     From MovServicios MC" _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cNumRecibo" _
            & " ) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where  Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "'" _
            & " And M.nMovFlag = 0 " & sqlAdd
            
            
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
            
   Else
        sTitulo1 = psTitulo & " - " & psUsuario & " - " & " DEL " & UCase(lsFec) & " - "
        SQL = " Select dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod, M.nMovNro nNumTran, Right(M.cMovNro,4) cCodUsu, O.cOpeDesc cNomOpe, O.cOpeCod, " _
            & " Ope.cCtaCod cCodCta, Ope.cDocNro cNumDoc, Ope.Monto nMonTran , SubString(M.cMovNro,18,2) cCodAge" _
            & " From Mov M" _
            & " Inner Join OpeTpo O On M.cOpeCod = O.cOpeCod" _
            & " Inner Join" _
            & "     (Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '000000002' cCtaCod, '' cDocNro     " _
            & "         From MovCompraVenta MC" _
            & "             Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & " Group By M.nMovNro, M.cOpeCod" _
            & " Union" _
            & " Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))) cCtaCod, Mc.cNroDoc" _
            & "     From MovOpeVarias MC " _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro " _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cNroDoc" _
            & " Union" _
            & "     Select Sum(nMovImporte) Monto, M.nMovNro, M.cOpeCod, '00000000' + LTrim(RTrim(Str(MC.nMoneda))) cCtaCod, 'Dest : ' + MC.cUsuDest cNroDoc" _
            & "     From MovHabDevCajero MC" _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cUsuDest"
            
        SQL = SQL & " Union" _
            & "     Select Sum(nMonto) Monto, M.nMovNro, M.cOpeCod, '00000000' + LTrim(RTrim(Str(MC.nMoneda))) cCtaCod, MC.cNumRecibo cNroDoc" _
            & "     From MovServicios MC" _
            & "     Inner Join Mov M On MC.nMovNro = M.nMovNro" _
            & "     Where Left(M.cMovNro,8) Between '" & Format(ldIni, gcFormatoMov) & "' And '" & Format(ldFin, gcFormatoMov) & "'" _
            & "     Group By M.nMovNro, M.cOpeCod, '00000000' + ltrim(rtrim(str(MC.nMoneda))), MC.cNumRecibo" _
            & " ) As Ope On Ope.nMovNro = M.nMovNro" _
            & " Where  Substring(Ope.cCtaCod,9,1) = '" & pnMoneda & "'" _
            & " And M.nMovFlag = 0 " & sqlAdd & "  And  Right(M.cMovNro,4) = '" & psUsuario & "'"
            
            If lsOrden = "1" Then
                SQL = SQL & " Order by nNumTran"
            ElseIf lsOrden = "2" Then
                SQL = SQL & " Order By O.cOpeCod "
            End If
            
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not RSVacio(rs) Then
        PonerSalto lsCadena
        lbBan = True
        
        lsCadena = lsCadena + CabeceraPagina(sTitulo1 & " OTROS", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
        lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
        
        sTEmpoOpe = rs!copecod
        sOperacionTempo = rs!cNomOpe
        nMonto = 0
        nCantidad = 0
        
        While Not rs.EOF
            
            If lsOrden = "2" Then 'por codigo de operacion
                If sTEmpoOpe <> rs!copecod Then
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & String(131, "-") & Chr(10)
                    End If
                    lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
                    If lsCheck <> "1" Then
                        lsCadena = lsCadena & Chr(10)
                    End If
                    lsNomOpe = rs!cNomOpe
                    sOperacionTempo = rs!cNomOpe
                    nMonto = rs!nMonTran
                    sTEmpoOpe = rs!copecod
                    nCantidad = 1
                Else
                    nMonto = nMonto + rs!nMonTran
                    nCantidad = nCantidad + 1
                End If
                nMontoT = nMontoT + rs!nMonTran
                nCantidadT = nCantidadT + 1
            End If
            
            If IsNull(rs!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = rs!cNumDoc
            End If
            
            If IsNull(rs!cCodCta) Then
                lsCodCta = " "
            Else
                lsCodCta = rs!cCodCta
            End If
            
            If IsNull(rs!cCodUsu) Then
                If IsNull(rs!cCodusurem) Then
                    lsCodUsu = ""
                Else
                    lsCodUsu = rs!cCodusurem
                End If
            Else
                lsCodUsu = rs!cCodUsu
            End If
            
            lsNomOpe = Mid(rs!cNomOpe, 1, 26)
            If lsCheck <> "1" Then
            lsCadena = lsCadena & Space(10 - Len(Str(rs!nNumTran))) & Str(rs!nNumTran) _
                                & Space(20 - Len(lsCodCta)) & lsCodCta _
                                & Space(2) & Trim(lsNomOpe) _
                                & Space(26 - Len(Trim(lsDoc)) + 20 - Len(Trim(lsNomOpe))) & Trim(lsDoc) _
                                & Space(16 - Len(Format(rs!nMonTran, "###,##0.00"))) & Format(rs!nMonTran, "###,##0.00") _
                                & Space(7 - Len(lsCodUsu)) & lsCodUsu _
                                & Space(7 - Len(rs!CCodAge)) & rs!CCodAge _
                                & Space(14 - Len(Format(rs!dFecTran, "hh:mm:ss AMPM"))) & Format(rs!dFecTran, "hh:mm:ss AMPM") & " [" & rs!copecod & "]" & Chr$(10)
            End If
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                lsCadena = lsCadena & Chr$(12)
                lsCadena = lsCadena & CabeceraPagina(sTitulo1 & " CREDITO", pnPagina, pnItem, pgsNomAge, pgsEmpresa, pgdFecSis, pnMoneda)
                lsCadena = lsCadena & Encabezado("Nro.Tran;10;Nro Cuenta;16;Tip.Ope.;24;Nro.Doc.;20;Monto;20;Usu.;12;Agencia;8;Hora;7; ;6;Cod.Ope.;8;", pnItem)
            End If
            
            rs.MoveNext
        Wend
        
        If lsOrden = "2" Then 'por codigo de operacion
            If lsCheck <> "1" Then
                lsCadena = lsCadena & String(131, "-") & Chr(10)
            End If
            lsCadena = lsCadena & Space(26) & "Total " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidad, 6, 0) & ")" & Space(1) & Space(13 - Len(Format(nMonto * 1, "###,##0.00"))) & Format(nMonto * 1, "###,##0.00") & Space(28) & " [" & sTEmpoOpe & "]" & Chr(10)
        End If
        
        'If lsOrden = "2" Then 'por codigo de operacion
        '    lsCadena = lsCadena & String(122, "=") & Chr(10)
        '    sOperacionTempo = ""
        '    lsCadena = lsCadena & Space(26) & "TOTAL " & sOperacionTempo & Space(7) & "(" & ImpreFormat(nCantidadT, 6, 0) & ")" & Space(1) & ImpreFormat(nMontoT, 10, 2) & Chr(10) & Chr(10)
        'End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If pbSalto Then lsCadena = lsCadena & Chr$(12)
        
    ProtocoloOperaciones = lsCadena
End Function

'
Public Function ComVtaME(ByVal pdIni As Date, ByVal pdFin As Date, psAge As String) As String
    Dim lsQryOpe As String
    Dim rsQryOpe As ADODB.Recordset
    Set rsQryOpe = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    Dim lsRTFCaj As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "'"
    End If
    
    
    lsRTFCaj = ""
    
    oCon.AbreConexion
    
    lsQryOpe = ""
    lsQryOpe = " SELECT dbo.FechaHoraMov(M.cMovNro) dFecTran, M.cOpeCod cCodOpe, M.nMovNro nNumTran, PE.cPersNombre cNomPers, CV.nMovImporte nMonTran, MTC.nMovTpoCambio nTipcambio, (CV.nMovImporte*MTC.nMovTpoCambio) AS nConv, Right(M.cMovNro,4) cCodUsu  " _
             & " FROM        Mov M" _
             & " INNER JOIN  MovCompraVenta CV ON CV.nMovNro = M.nMovNro" _
             & " INNER JOIN  Persona PE ON PE.cPersCod = CV.cPersCod" _
             & " INNER JOIN  MovTpoCambio MTC ON MTC.nMovNro = M.nMovNro" _
             & " WHERE       M.cOpeCod IN ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMEVenta & "','" & gOpeCajeroMECompraEsp & "', '" & gOpeCajeroMEVentaEsp & "') AND" _
             & " LEFT (M.cMovNro, 8) >= '" & Format(pdIni, "YYYYMMdd") & "' and LEFT (M.cMovNro, 8) <='" & Format(pdFin, "YYYYMMdd") & "' AND M.nMovFlag = " & MovFlag.gMovFlagVigente _
             & "  " & sqlAdd _
             & " ORDER BY M.cOpeCod"
            
            'Left(M.cMovNro,8) BETWEEN '" & Format(pdIni, gsFormatoMovFecha) & "' AND '" & Format(DateAdd("d", 1, pdFin), gsFormatoMovFecha) & "'
             
    If rsQryOpe.State = adStateOpen Then rsQryOpe.Close
    Set rsQryOpe = oCon.CargaRecordSet(lsQryOpe)
    If RSVacio(rsQryOpe) Then
       lsRTFCaj = ""
    Else
       lsRTFCaj = PrtComVta(rsQryOpe, pdIni, pdFin)
    End If
    ComVtaME = lsRTFCaj
End Function

Private Function PrtComVta(prQryDat As ADODB.Recordset, pdIni As Date, pdFin As Date) As String

    Dim lnCarLin As Integer, i As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnSumMon As Currency, lnSumCon As Currency
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lsCodCta As String
    Dim lnPosDer As Integer, lsCodMon As String
    Dim lsRTFCaj As String
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 112
    lsTitRp1 = "LISTADO DIARIO MONEDA EXTRANJERA " & lsAgeCod
    If pdIni = pdFin Then
        lsTitRp2 = "DEL " & Format(CDate(pdIni), "dd mmmm yyyy")
    Else
        lsTitRp2 = "DEL " & pdIni & " AL " & pdFin
    End If
    lnCntPag = 0
    With prQryDat
         Do While Not .EOF
            If !cCodOpe <> lsCodMon Then
               If !cCodOpe = gOpeCajeroMECompra Then
                    lsSubTit = " - COMPRA"
               ElseIf !cCodOpe = gOpeCajeroMECompraEsp Then
                    lsSubTit = " - COMPRA ESPECIAL"
               ElseIf !cCodOpe = gOpeCajeroMEVenta Then
                    lsSubTit = " - VENTA"
               ElseIf !cCodOpe = gOpeCajeroMEVentaEsp Then
                    lsSubTit = " - VENTA ESPECIAL"
               End If
               If lsNumPag <> "" Then
                  lsRTFCaj = lsRTFCaj & String(lnCarLin, "-") & Chr$(10)
                  lsRTFCaj = lsRTFCaj & "Total" & String(37, " ")
                  lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumMon)), 12, True, 9, 2) & String(13, " ")
                  lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumCon)), 12, True, 9, 2) & Chr$(10)
                  lsRTFCaj = lsRTFCaj & String(lnCarLin, "=") & Chr$(12)
                  lsCodMon = !cCodOpe
                  lnSumMon = 0: lnSumCon = 0
               End If
               lsCodMon = !cCodOpe
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Then
               If lsNumPag <> "" Then
                  If lnSumMon <> 0 Or lnSumCon <> 0 Then
                     lsRTFCaj = lsRTFCaj & String(lnCarLin, "-") & Chr$(10)
                     lsRTFCaj = lsRTFCaj & "Total" & String(37, " ")
                     lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumMon)), 12, True, 9, 2) & String(13, " ")
                     lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumCon)), 12, True, 9, 2) & Chr$(10) & Chr$(12)
                  End If
                  lsCodMon = !cCodOpe
                  lnSumMon = 0: lnSumCon = 0
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsMoneda = "SOLES"
               
               lsRTFCaj = lsRTFCaj & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1 & lsSubTit, lsTitRp2, String(7, " "), lsNumPag) & Chr$(10)
               
               lsRTFCaj = lsRTFCaj & String(lnCarLin, "-") & Chr$(10)
               lsRTFCaj = lsRTFCaj & "NUMERO "
               lsRTFCaj = lsRTFCaj & String(12, " ") & "C L I E N T E" & String(11, " ") & "    "
               lsRTFCaj = lsRTFCaj & "DOLARES" & "    "
               lsRTFCaj = lsRTFCaj & "CAMBIO" & "          "
               lsRTFCaj = lsRTFCaj & "SOLES" & "  "
               lsRTFCaj = lsRTFCaj & "USUARIO" & "     "
               lsRTFCaj = lsRTFCaj & "FECHA / HORA" & Chr$(10)
               lsRTFCaj = lsRTFCaj & String(lnCarLin, "-") & Chr$(10)
               lnLinPag = 9
            End If
            lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(!nNumTran)), 6, False, 6, 0) & "   "
            
            If Len(Trim(ImpreCarEsp(PstaNombre(!cNomPers, False)))) >= 30 Then
               lsRTFCaj = lsRTFCaj & Left(ImpreCarEsp(PstaNombre(!cNomPers, False)), 30) & "   "
            Else
               lsRTFCaj = lsRTFCaj & FillText(Trim(ImpreCarEsp(PstaNombre(!cNomPers, False))), 30, " ") & "   "
            End If
            
            lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(!nMonTran)), 12, True, 9, 2) & "   "
            lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(!nTipCambio)), 7, True, 3, 3) & "   "
            lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(!nConv)), 12, True, 9, 2) & "   "
            lsRTFCaj = lsRTFCaj & !cCodUsu & "   "
            lsRTFCaj = lsRTFCaj & UCase(Format(!dFecTran, "dd/MM/YYYY") & " " & Format(!dFecTran, "hh:mm:ss AMPM")) & Chr$(10)
            lnSumMon = lnSumMon + !nMonTran
            lnSumCon = lnSumCon + !nConv
            lnLinPag = lnLinPag + 1
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFCaj = lsRTFCaj & String(lnCarLin, "-") & Chr$(10)
       lsRTFCaj = lsRTFCaj & "Total" & String(37, " ")
       lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumMon)), 12, True, 9, 2) & String(13, " ")
       lsRTFCaj = lsRTFCaj & JDNum(Trim(Str(lnSumCon)), 12, True, 9, 2) & Chr$(10)
       lsRTFCaj = lsRTFCaj & String(lnCarLin, "=") & Chr$(10) & Chr$(12)
    End If
    PrtComVta = lsRTFCaj
End Function


Public Function InfCajaGeneral(Optional ByVal pdFec As Date) As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim pR As String
    Dim BLMS As String
    Dim BLMD As String
    Dim Total As Currency
    Dim TOTALD As Currency
    Dim NegritaOn As String
    Dim NegritaOff As String
    Dim MMS(50) As String
    Dim MMD(50) As String
    Dim BMS(50) As String
    Dim BMD(50) As String
    Dim Sql1 As String
    Dim rs1 As New ADODB.Recordset
    Dim i As Integer
    Dim Check As Boolean
    Dim Check2 As Boolean
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim sCad As String
    Dim lnCarLin As Integer
    Dim bHoy As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta



    oCon.AbreConexion
    
    bHoy = False
    If DateDiff("d", pdFec, gdFecSis) = 0 Then bHoy = True
    If bHoy Then
        'Verifica que se haya generado el cierre de dia o el cierre de mes para generar
        'los mensajes de prevención correctos
        If CierreRealizado Then
            If VerificaDiaHabil(gdFecSis, 3) Then
                If CierreRealizado(2) Then
                    lsTitRp2 = "USO DE CONTABILIDAD Y CAJA GENERAL"
                Else
                    lsTitRp2 = "NO VALIDO (ANTES DEL CIERRE DEL MES)"
                End If
            Else
                lsTitRp2 = "USO DE CONTABILIDAD Y CAJA GENERAL"
            End If
        Else
            lsTitRp2 = "NO VALIDO (ANTES DEL CIERRE DEL DIA)"
        End If
    Else
        lsTitRp2 = UCase(Format$(pdFec, "dd mmmm yyyy"))
    End If
    
    'Obtiene el nombre de las monedas en soles
    Sql1 = "Select nEfectivoValor from Efectivo Where cEfectivoCod Like '1M%' Order By nEfectivoValor Desc"
    Set rs1 = oCon.CargaRecordSet(Sql1)
    i = 1
    If rs1.EOF And rs1.BOF Then
     MMS(1) = 0
    Else
        Do While Not rs1.EOF
           MMS(i) = Trim(rs1!nEfectivoValor)
           rs1.MoveNext
           i = i + 1
        Loop
    End If
    rs1.Close
    'Obtiene el nombre de las Monedas Dolares
    Sql1 = "Select nEfectivoValor from Efectivo Where cEfectivoCod Like '2M%' Order By nEfectivoValor Desc"
    Set rs1 = oCon.CargaRecordSet(Sql1)
    i = 1
    If RSVacio(rs1) Then
     MMD(1) = 0
    Else
        Do While Not rs1.EOF
           MMD(i) = Trim(rs1!nEfectivoValor)
           rs1.MoveNext
           i = i + 1
        Loop
    End If
    rs1.Close
    
    'Obtiene el nombre de las Billetes Soles
    Sql1 = "Select nEfectivoValor from Efectivo Where cEfectivoCod Like '1B%' Order By nEfectivoValor Desc"
    Set rs1 = oCon.CargaRecordSet(Sql1)
    i = 1
    If RSVacio(rs1) Then
     BMS(1) = 0
    Else
        Do While Not rs1.EOF
           BMS(i) = Trim(rs1!nEfectivoValor)
           rs1.MoveNext
           i = i + 1
        Loop
    End If
    rs1.Close
    
    'Obtiene el nombre de las Billetes Dolares
    Sql1 = "Select nEfectivoValor from Efectivo Where cEfectivoCod Like '2B%' Order By nEfectivoValor Desc"
    Set rs1 = oCon.CargaRecordSet(Sql1)
    i = 1
    If RSVacio(rs1) Then
     BMD(1) = 0
    Else
        Do While Not rs1.EOF
           BMD(i) = Trim(rs1!nEfectivoValor)
           rs1.MoveNext
           i = i + 1
        Loop
    End If
    rs1.Close
    
    NegritaOn = gPrnBoldON
    NegritaOff = gPrnBoldOFF
    TB = 0
    TBD = 0
    lnCarLin = 122
    lsTitRp1 = "INFORME PARA CAJA GENERAL"
    sCad = sCad & CabeRepo("", "", lnCarLin, "", lsTitRp1, lsTitRp2, "", "") & Chr$(10)
    
    '*** AHORROS SOLES y DOLARES*****************************************************
    sCad = sCad & NegritaOn & "I ." & Space(2) & "AHORROS" & NegritaOff & Chr$(10)
    sCad = sCad & NegritaOn & Space(3) & "1.1  Moneda Nacional  " & Space(10) & "#" & Space(15) & "Saldos"
    sCad = sCad & Space(5) & "1.2  Moneda Extranjera" & Space(13) & "#" & Space(15) & "Saldos" & NegritaOff & Chr$(10)
    sCad = sCad & Ahorros(pdFec)
    sCad = sCad & Chr$(10)
    
      '****CREDITOS SOLES Y DOLARES ****
    sCad = sCad & NegritaOn & "II. " & "CREDITOS" & NegritaOff & Chr$(10)
    sCad = sCad & NegritaOn & Space(3) & "2.1  Moneda Nacional  " & Space(10) & "#" & Space(15) & "Saldos"
    sCad = sCad & Space(5) & "2.2  Moneda Extranjera" & Space(13) & "#" & Space(15) & "Saldos" & NegritaOff & Chr$(10)
    sCad = sCad & Creditos(pdFec)
    sCad = sCad & Chr$(10)
    
    '********** PRENDARIO *************************
    sCad = sCad & NegritaOn & "III." & Space(2) & "CREDITO PRENDARIO" & Space(12) & "#" & Space(10) & "Gramos" & Space(10) & "Saldos" + NegritaOff & Chr$(10)
     pR = DP("V")
    sCad = sCad & Space(8) & "Cred. Vigentes    " & Space(2) & JDNum(pR, 10, False, 8, 0) & Space(4) & JDNum(NumP, 12, False, 12, 2) & Space(4) & JDNum(SalP, 13, True, 11, 2) & Chr$(10)
     pR = DP("A")
    sCad = sCad & Space(8) & "Cred. Adjudicados " & Space(2) & JDNum(pR, 10, False, 8, 0) & Space(4) & JDNum(NumP, 12, False, 12, 2) & Space(4) & JDNum(SalP, 13, True, 11, 2) & Chr$(10)
     pR = DP("D")
    sCad = sCad & Space(8) & "Cred. Diferidos   " & Space(2) & JDNum(pR, 10, False, 8, 0) & Space(4) & JDNum(NumP, 12, False, 12, 2) & Space(4) & JDNum(SalP, 13, True, 11, 2) & Chr$(10)
         
    '************CAJA ******************
    sCad = sCad & NegritaOn & "IV. " & Space(3) & "CAJA" + NegritaOff & Chr$(10)
    'MONEDA NACIONAL
    sCad = sCad & NegritaOn & Space(3) & "4.1  Moneda Nacional" & Space(43) & "4.2  Moneda Extranjera" + NegritaOff & Chr$(10)
    'BILLETES SOLES
    sCad = sCad & Space(8) & "Billetes" & Space(53) & "Billetes" & Chr$(10)

    For i = 1 To 50
     If BMS(i) <> "" And BMD(i) <> "" Then
     BLMS = BL("1", BMS(i), "B", pdFec)
     BLMD = BL("2", BMD(i), "B", pdFec)
     
     sCad = sCad & Space(13) & JDNum(BMS(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CA), 9, True, 9, 0) & Space(4) & JDNum(BLMS, 13, True, 11, 2) & Space(15) & JDNum(BMD(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CAE), 9, True, 9, 0) & Space(4) & JDNum(BLMD, 13, True, 11, 2) & Chr$(10)
    Else
        If BMS(i) = "" And BMD(i) = "" Then
         Exit For
        End If
        If BMS(i) = "" Then
        BLMD = BL("2", BMD(i), "B", pdFec)
        sCad = sCad & Space(72) & JDNum(BMD(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CAE), 9, True, 9, 0) & Space(4) & JDNum(BLMD, 13, True, 11, 2) & Chr$(10)
        End If
        If BMD(i) = "" Then
        BLMS = BL("1", BMS(i), "B", pdFec)
        sCad = sCad & Space(13) & JDNum(BMS(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CA), 9, True, 9, 0) & Space(4) & JDNum(BLMS, 13, True, 11, 2)
        End If
    End If
    Next
    sCad = sCad & NegritaOn & Space(13) & "Total Billetes " & Space(16) & JDNum(Str(TB), 13, True, 11, 2) & Space(13) & "Total Billetes " & Space(18) & JDNum(Str(TBD), 13, True, 11, 2) & NegritaOff & Chr$(10) & Chr$(10)
    Total = Total + TB
    TOTALD = TOTALD + TBD
    TB = 0
    TBD = 0
    
    
    sCad = sCad & Space(8) & "Monedas " & Space(53) & "Monedas " & Chr$(10)
    For i = 1 To 50
        If MMS(i) <> "" And MMD(i) <> "" Then
          BLMS = BL("1", MMS(i), "M", pdFec)
          BLMD = BL("2", MMD(i), "M", pdFec)
          sCad = sCad & Space(13) & JDNum(MMS(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CA), 9, True, 9, 0) & Space(4) & JDNum(BLMS, 13, True, 11, 2) & Space(15) & JDNum(MMD(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CAE), 9, True, 9, 0) & Space(4) & JDNum(BLMD, 13, True, 11, 2) & Chr$(10)
         Else
            If MMS(i) = "" And MMD(i) = "" Then
              Exit For
            End If
            If MMS(i) = "" Then
            BLMD = BL("2", MMD(i), "M", pdFec)
            sCad = sCad & Space(72) & JDNum(MMD(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CAE), 9, True, 9, 0) & Space(4) & JDNum(BLMD, 13, True, 11, 2) & Chr$(10)
            End If
            If MMD(i) = "" Then
            BLMS = BL("1", MMS(i), "M", pdFec)
            sCad = sCad & Space(13) & JDNum(MMS(i), 8, True, 6, 2) & Space(10) & JDNum(Str(CA), 9, True, 9, 0) & Space(4) & JDNum(BLMS, 13, True, 11, 2) & Chr$(10)
            End If
        End If
    Next
    sCad = sCad & NegritaOn & Space(14) & "Total Monedas " & Space(16) & JDNum(Str(TB), 13, True, 11, 2) & Space(14) & "Total Monedas " & Space(18) & JDNum(Str(TBD), 13, True, 11, 2) & NegritaOff & Chr$(10) & Chr$(10)
    Total = Total + TB
    TOTALD = TOTALD + TBD
    sCad = sCad & NegritaOn & Space(14) & "         TOTAL" & Space(16) & JDNum(Str(Total), 13, True, 11, 2) & Space(14) & "        TOTAL " & Space(16) & JDNum(Str(TOTALD), 13, True, 11, 2) + NegritaOff & Chr$(10)
    sCad = sCad & Chr$(12)
    InfCajaGeneral = sCad
End Function


Public Function SaldosDiarios(ByVal dFec As Date, ByVal sAgeCod As String, _
        ByVal sAgeNom As String, ByVal bHoy As Boolean, ByVal sNomCmac As String) As String
    Dim Cad As String
    Dim NON As String
    Dim NOF As String
    
    NON = gPrnBoldON
    NOF = gPrnBoldOFF
    Cad = Cad & NON & sNomCmac & Chr$(10)
    Cad = Cad & sAgeNom & Chr$(10) & Chr$(10)
    Cad = Cad & "A:        CONTABILIDAD -  " & ldFecRepo & Chr$(10)
    Cad = Cad & "DE:       " + UCase(sAgeNom) & Space(20) & "Fecha : " + UCase(ArmaFecha(dFec)) + NOF & Chr$(10) & Chr$(10)
    Cad = Cad & String(80, "=") & Chr$(10)
    Cad = Cad & NON & Space(25) & "SALDOS DIARIOS DE MONEDA EXTRANJERA" + NOF & Chr$(10)
    Cad = Cad & String(80, "=") & Chr$(10)
    Cad = Cad & Space(35) & "EN U.S DOLARES " & Chr$(10) & Chr$(10)
    Cad = Cad & "1. CAJA             " & Space(40) & JDNum(DCaja(dFec, sAgeCod), 13, True, 11, 2) & Chr$(10)
    Cad = Cad & "2. COLOCACIONES     " & Space(40) & JDNum(DColocaciones(dFec, sAgeCod, , bHoy), 13, True, 11, 2) & Chr$(10)
    Cad = Cad & "3. DEPOSITOS DE AHORRO" & Space(38) & JDNum(DAhorro(dFec, sAgeCod, , bHoy), 13, True, 11, 2) & Chr$(10)
    Cad = Cad & "4. DEPOSITOS A PLAZO" & Space(40) & JDNum(DPlazo(dFec, sAgeCod, , bHoy), 13, True, 11, 2) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
    Cad = Cad & String(80, "=") & Chr$(10)
    Cad = Cad & Space(25) + NON & "REPORTE DE COMPRA Y VENTA DE DOLARES" + NOF & Chr$(10)
    Cad = Cad & String(80, "=") & Chr$(10)

    Cad = Cad & NON & "REPORTE DE COMPRAS:" + NOF & Chr$(10)
    Cad = Cad & Chr$(10)
    Cad = Cad & Space(5) & "CANTIDAD COMPRADA EN DOLARES" & Space(27) & JDNum(CompraVenta("C", dFec), 13, True, 11, 2) & Chr$(10)
    Cad = Cad & Space(5) & "IMPORTE PAGADO EN SOLES     " & Space(27) & JDNum(ImpSoles, 13, True, 11, 2) & Chr$(10) & Chr$(10)

    Cad = Cad & NON & "REPORTE DE VENTAS:" + NOF & Chr$(10)
    Cad = Cad & Chr$(10)
    Cad = Cad & Space(5) & "CANTIDAD VENDIDA EN DOLARES " & Space(27) & JDNum(CompraVenta("V", dFec), 13, True, 11, 2) & Chr$(10)
    Cad = Cad & Space(5) & "IMPORTE COBRADO EN SOLES    " & Space(27) & JDNum(ImpSoles, 13, True, 11, 2) & Chr$(10) & Chr$(10) & Chr$(10)
    Cad = Cad & NON & "TIPOS DE CAMBIO DEL DIA" + NOF & Chr$(10)
    Cad = Cad & Chr$(10)
    Cad = Cad & "TIPO DE CAMBIO DE COMPRA    " & Chr$(10)
    TCambio dFec
    For i = 1 To k - 1
        Cad = Cad & Space(5) & "Hora:  " + DHora(i) & Space(40) & JDNum(TC(i), 8, True, 6, 2) & JDNum(TCE(i), 8, True, 6, 2) & Chr$(10)
    Next
        Cad = Cad & "TIPO DE CAMBIO DE VENTA     " & Chr$(10)
    For i = 1 To k - 1
        Cad = Cad & Space(5) & "Hora:  " + DHora(i) & Space(40) & JDNum(TV(i), 8, True, 6, 2) & JDNum(TVE(i), 8, True, 6, 2) & Chr$(10)
    Next
    SaldosDiarios = Cad
End Function


Public Function CierreRealizado(Optional pnTipo As Integer = 1) As Boolean
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    If pnTipo = 1 Then
        CierreRealizado = IIf(oConst.LeeConstSistema(gConstSistCierreSistema) = oConst.LeeConstSistema(gConstSistFechaSistema), True, False)
    ElseIf pnTipo = 2 Then
        CierreRealizado = IIf(oConst.LeeConstSistema(gConstSistCierreMesNegocio) = oConst.LeeConstSistema(gConstSistFechaSistema), True, False)
    End If
End Function

Private Function VerificaDiaHabil(ByVal pdFecha As Date, pnTipo As Integer) As Boolean
    Dim lsFchPro As String
    Dim ldFchPro As Date
    Dim lnMes As Integer
    Dim ldAno As Date
    ldAno = pdFecha
    lnMes = Month(pdFecha)
    If lnMes = 12 Then
        lnMes = 0
        ldAno = DateAdd("yyyy", 1, ldAno)
    End If
Select Case pnTipo
    Case 1
    '   Busca Primer día del mes
        lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
        ldFchPro = CDate(lsFchPro)
    Case 2
    '   Busca día anterior al último día del mes para procesar descuento por Inactivas
        lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
        ldFchPro = DateAdd("d", -2, CDate(lsFchPro))
    Case 3
    '   Busca ultimo día del mes para procesar Cierre de Mes
        lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
        ldFchPro = DateAdd("d", -1, CDate(lsFchPro))
End Select
If pdFecha = ldFchPro Then
    VerificaDiaHabil = True
Else
    VerificaDiaHabil = False
End If
End Function


Private Function Ahorros(ByVal pdFec As Date) As String
    Dim TotalAMN As Double, TotalAME As Double
    Dim sCad As String
    Dim TcMN As Double, TcME As Double ' TOTAL CHEQUES
    Dim NegritaOn As String
    Dim NegritaOff As String
    NegritaOn = gPrnBoldON
    NegritaOff = gPrnBoldOFF
    Dim RegMN As Long, RegME As Long
    Dim SaldCmact As String
    Dim RegCmacS As Integer

    'Inicializamos las variables de totales
    TotalAMN = 0: TotalAME = 0
    TcMN = 0: TcME = 0
    RegMN = 0: RegME = 0
    sCad = ""
    '************ Depositos de Ahorro y CPE/Recursos Propios ****************'
    'Cajas Municipales : Depósitos de Ahorros
    SaldCmact = SaldosInstFinanc(gMonedaNacional, pdFec, Producto.gCapAhorros, gPersonaJurCFLCMAC)
    RegCmacS = RegCmacStemp
    sCad = sCad & Space(8) & "CMACS Dep. Ahorros  " & Space(5) & JDNum(Str(RegCmacS), 6, False, 6, 0) & Space(7) & JDNum(SaldCmact, 13, True, 11, 2)
    SaldCmact = SaldosInstFinanc(gMonedaExtranjera, pdFec, Producto.gCapAhorros, gPersonaJurCFLCMAC)
    RegCmacS = RegCmacStemp
    sCad = sCad & Space(8) & "CMACS Dep. Ahorros  " & Space(5) & JDNum(Str(RegCmacS), 9, False, 9, 0) & Space(8) & JDNum(SaldCmact, 13, True, 11, 2) & Chr$(10)
    'Cajas Municipales : Depósitos a Plazo
    SaldCmact = SaldosInstFinanc(gMonedaNacional, pdFec, Producto.gCapPlazoFijo, gPersonaJurCFLCMAC)
    RegCmacS = RegCmacStemp
    sCad = sCad & Space(8) & "CMACS Dep. a Plazo  " & Space(5) & JDNum(Str(RegCmacS), 6, False, 6, 0) & Space(7) & JDNum(SaldCmact, 13, True, 11, 2)
    SaldCmact = SaldosInstFinanc(gMonedaExtranjera, pdFec, Producto.gCapPlazoFijo, gPersonaJurCFLCMAC)
    RegCmacS = RegCmacStemp
    sCad = sCad & Space(8) & "CMACS Dep. a Plazo  " & Space(5) & JDNum(Str(RegCmacS), 9, False, 9, 0) & Space(8) & JDNum(SaldCmact, 13, True, 11, 2) & Chr$(10)
    'Otras Instituciones Financieras : Depósitos de Ahorros
    RegCmacS = 0
    SaldCmact = SaldosInstFinanc(gMonedaNacional, pdFec, Producto.gCapAhorros, gPersonaJurCFLCRAC)
    sCad = sCad & Space(8) & "O.I.F Dep. Ahorros  " & Space(5) & JDNum(Str(RegCmacS), 6, False, 6, 0) & Space(7) & JDNum(SaldCmact, 13, True, 11, 2)
    SaldCmact = SaldosInstFinanc(gMonedaExtranjera, pdFec, Producto.gCapAhorros, gPersonaJurCFLCRAC)
    sCad = sCad & Space(8) & "O.I.F Dep. Ahorros  " & Space(5) & JDNum(Str(RegCmacS), 9, False, 9, 0) & Space(8) & JDNum(SaldCmact, 13, True, 11, 2) & Chr$(10)
    'Otras Instituciones Financieras : Depósitos de Ahorros
    SaldCmact = SaldosInstFinanc(gMonedaNacional, pdFec, Producto.gCapPlazoFijo, gPersonaJurCFLCRAC)
    sCad = sCad & Space(8) & "O.I.F Dep. a Plazo  " & Space(5) & JDNum(Str(RegCmacS), 6, False, 6, 0) & Space(7) & JDNum(SaldCmact, 13, True, 11, 2)
    SaldCmact = SaldosInstFinanc(gMonedaExtranjera, pdFec, Producto.gCapPlazoFijo, gPersonaJurCFLCRAC)
    sCad = sCad & Space(8) & "O.I.F Dep. a Plazo  " & Space(5) & JDNum(Str(RegCmacS), 9, False, 9, 0) & Space(8) & JDNum(SaldCmact, 13, True, 11, 2) & Chr$(10)
    DA Producto.gCapAhorros, gMonedaNacional, pdFec
    TotalAMN = TotalAMN + CCur(SalA)
    RegMN = RegMN + CLng(NumA)
    sCad = sCad & Space(8) & "Depósitos de Ahorro " & Space(5) & JDNum(NumA, 6, False, 6, 0) & Space(7) & JDNum(SalA, 13, True, 10, 2)
    DA Producto.gCapAhorros, gMonedaExtranjera, pdFec
    RegME = RegME + CLng(NumA)
    TotalAME = TotalAME + CCur(SalA)
    sCad = sCad & Space(8) & "Depósitos de Ahorro " & Space(5) & JDNum(NumA, 9, False, 9, 0) & Space(8) & JDNum(SalA, 13, True, 11, 2) & Chr$(10)
    DA Producto.gCapPlazoFijo, gMonedaNacional, pdFec
    TotalAMN = TotalAMN + CCur(SalA)
    RegMN = RegMN + CLng(NumA)
    sCad = sCad & Space(8) & "Depósitos a Plazo   " & Space(5) & JDNum(NumA, 6, False, 6, 0) & Space(7) & JDNum(SalA, 13, True, 10, 2)
    DA Producto.gCapPlazoFijo, gMonedaExtranjera, pdFec
    TotalAME = TotalAME + CCur(SalA)
    RegME = RegME + CLng(NumA)
    sCad = sCad & Space(8) & "Depósitos a Plazo   " & Space(5) & JDNum(NumA, 9, False, 9, 0) & Space(8) & JDNum(SalA, 13, True, 11, 2) & Chr$(10)
    DA "Cheques", gMonedaNacional, pdFec, Producto.gCapAhorros
    RegMN = RegMN + CLng(NumA)
    TcMN = TcMN + CCur(SalA)
    sCad = sCad & Space(8) & "AC/Cheques Valorización" & Space(2) & JDNum(NumA, 6, False, 6, 0) & Space(7) & JDNum(SalA, 13, True, 10, 2)
    DA "Cheques", gMonedaExtranjera, pdFec, Producto.gCapAhorros
    RegME = RegME + CLng(NumA)
    TcME = TcME + CCur(SalA)
    sCad = sCad & Space(8) & "AC/Cheques Valorización" & Space(2) & JDNum(NumA, 9, False, 9, 0) & Space(8) & JDNum(SalA, 13, True, 11, 2) & Chr$(10)
    DA "Cheques", gMonedaNacional, pdFec, Producto.gCapPlazoFijo
    TcMN = TcMN + CCur(SalA)
    RegMN = RegMN + CLng(NumA)
    sCad = sCad & Space(8) & "PF/Cheques Valorización" & Space(2) & JDNum(NumA, 6, False, 6, 0) & Space(7) & JDNum(SalA, 13, True, 10, 2)
    DA "Cheques", gMonedaExtranjera, pdFec, Producto.gCapPlazoFijo
    TcME = TcME + CCur(SalA)
    RegME = RegME + CLng(NumA)
    sCad = sCad & Space(8) & "PF/Cheques Valorización" & Space(2) & JDNum(NumA, 9, False, 9, 0) & Space(8) & JDNum(SalA, 13, False, 11, 2) & Chr$(10)
    DA "Cheques", gMonedaNacional, pdFec, Producto.gCapCTS
    TcMN = TcMN + CCur(SalA)
    RegMN = RegMN + CLng(NumA)
    sCad = sCad & Space(8) & "CTS/Cheques Valorización" & Space(1) & JDNum(NumA, 6, False, 6, 0) & Space(7) & JDNum(SalA, 13, True, 10, 2)
    DA "Cheques", gMonedaExtranjera, pdFec, Producto.gCapCTS
    TcME = TcME + CCur(SalA)
    RegME = RegME + CLng(NumA)
    sCad = sCad & Space(8) & "CTS/Cheques Valorización" & Space(1) & JDNum(NumA, 9, False, 9, 0) & Space(8) & JDNum(SalA, 13, False, 11, 2) & Chr$(10)
    TotalAMN = TotalAMN - TcMN
    TotalAME = TotalAME - TcME
    sCad = sCad & NegritaOn & Space(15) & "Total     " & Space(4) & JDNum(Str(RegMN), 10, False, 10, 0) & Space(7) & JDNum(Str(TotalAMN), 13, True, 10, 2)
    sCad = sCad & Space(20) & "Total     " & Space(3) & JDNum(Str(RegME), 9, False, 9, 0) & Space(8) & JDNum(Str(TotalAME), 13, True, 11, 2) & Chr$(10)
    
    'Devuelve la cadena generada
    Ahorros = sCad
End Function


Private Function Creditos(ByVal pdFec As Date) As String
    Dim sCad As String
    Dim TotalCMN As Double, TotalCME As Double
    Dim stCPEMN As Double, stCPEME As Double
    Dim stPPMN As Double, stPPME As Double
    Dim NegritaOn As String
    Dim NegritaOff As String
    Dim RegCMN As Long, RegCME As Long
    Dim stRegCPEMN As Long, stRegCPEME As Long
    Dim stRegPPMN As Long, stRegPPME As Long
    
    NegritaOn = gPrnBoldON
    NegritaOff = gPrnBoldOFF
    
    stCPEMN = 0: stCPEME = 0
    stPPMN = 0: stPPME = 0
    TotalCMN = 0: TotalCME = 0
    stRegCPEMN = 0: stRegCPEME = 0
    stRegPPMN = 0: stRegPPME = 0
    RegCMN = 0: RegCME = 0
    '************ CPE/Recursos Propios ****************'
    DC "2", "01", "1", pdFec      'CPE/Rec Prop
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegCPEMN = stRegCPEMN + CLng(NumC)
    stCPEMN = stCPEMN + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Recursos Propios" & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "2", "01", "2", pdFec  'CPE/Rec Prop
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegCPEME = stRegCPEME + CLng(NumC)
    stCPEME = stCPEME + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Recursos Propios" & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
      
    DC "2", "02", "1", pdFec  ' CPE/Foncode
    TotalCMN = TotalCMN + CCur(SalC)
    stRegCPEMN = stRegCPEMN + CLng(NumC)
    stCPEMN = stCPEMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    sCad = sCad & Space(8) & "CPE/Foncodes        " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "2", "02", "2", pdFec ' CPE/Foncode
    TotalCME = TotalCME + CCur(SalC)
    stRegCPEME = stRegCPEME + CLng(NumC)
    stCPEME = stCPEME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    sCad = sCad & Space(8) & "CPE/Foncodes        " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
      
    DC "2','4", "02', '03','05','07", "1", pdFec  ' CPE/Cofide -- CPE/Cofide Microglobal / CPE/PROPEM
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegCPEMN = stRegCPEMN + CLng(NumC)
    stCPEMN = stCPEMN + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Cofide          " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "2','4", "02', '03','05','07", "2", pdFec   ' CPE/Cofide -- CPE/Cofide Microglobal
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegCPEME = stRegCPEME + CLng(NumC)
    stCPEME = stCPEME + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Cofide          " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
     
    DC "2", "04", "1", pdFec   ' CPE/BID
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegCPEMN = stRegCPEMN + CLng(NumC)
    stCPEMN = stCPEMN + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/BID             " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "2", "04", "2", pdFec  ' CPE/BID
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegCPEME = stRegCPEME + CLng(NumC)
    stCPEME = stCPEME + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/BID             " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
      
    DC "2", "06", "1", pdFec  'CPE/Habitat  ' Nuevo
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegCPEMN = stRegCPEMN + CLng(NumC)
    stCPEMN = stCPEMN + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Habitat Prod.   " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "2", "06", "2", pdFec  'CPE/Habitat  ' Nuevo
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegCPEME = stRegCPEME + CLng(NumC)
    stCPEME = stCPEME + CCur(SalC)
    sCad = sCad & Space(8) & "CPE/Habitat Prod.   " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, False, 11, 2) & Chr$(10)
    
    'Sub Total CPE
    sCad = sCad & Space(15) & NegritaOn & "Sub Total" & Space(9) & JDNum(Str(stRegCPEMN), 6, False, 6, 0) & Space(7) & JDNum(Str(stCPEMN), 13, True, 11, 2)
    sCad = sCad & Space(15) & "Sub Total" & Space(8) & JDNum(Str(stRegCPEME), 10, False, 10, 0) & Space(8) & JDNum(Str(stCPEME), 13, True, 11, 2) + NegritaOff & Chr$(10)
    
    DC "3", "01", "1", pdFec   ' PP/Dscto Planill
    TotalCMN = TotalCMN + CCur(SalC)
    stPPMN = stPPMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegPPMN = stRegPPMN + CLng(NumC)
    sCad = sCad & Space(8) & "PP/Desc. Planilla   " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "3", "01", "2", pdFec ' PP/Dscto Planill
    TotalCME = TotalCME + CCur(SalC)
    stPPME = stPPME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegPPME = stRegPPME + CLng(NumC)
    sCad = sCad & Space(8) & "PP/Desc. Planilla   " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
    
    DC "3", "04", "1", pdFec ' PP/Usos Divers
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegPPMN = stRegPPMN + CLng(NumC)
    stPPMN = stPPMN + CCur(SalC)
    sCad = sCad & Space(8) & "PP/Usos Diversos    " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "3", "04", "2", pdFec ' PP/Usos Divers
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegPPME = stRegPPME + CLng(NumC)
    stPPME = stPPME + CCur(SalC)
    sCad = sCad & Space(8) & "PP/Usos Diversos    " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
     
    DC "3", "02", "1", pdFec  ' PP/PF
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegPPMN = stRegPPMN + CLng(NumC)
    stPPMN = stPPMN + CCur(SalC)
    sCad = sCad & Space(8) & "PP/Plazo Fijo       " & Space(5) & JDNum(NumC, 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "3", "02", "2", pdFec  ' PP/PF
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegPPME = stRegPPME + CLng(NumC)
    stPPME = stPPME + CCur(SalC)
    sCad = sCad & Space(8) & "PP/Plazo Fijo       " & Space(5) & JDNum(NumC, 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
    
    DC "3", "03", "1", pdFec ' PP/CTS
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegPPMN = stRegPPMN + CLng(NumC)
    stPPMN = stPPMN + CCur(SalC)
    sCad = sCad & Space(8) & "PP/CTS              " & Space(5) & JDNum(Str(NumC), 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "3", "03", "2", pdFec ' PP/CTS
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegPPME = stRegPPME + CLng(NumC)
    stPPME = stPPME + CCur(SalC)
    sCad = sCad & Space(8) & "PP/CTS              " & Space(5) & JDNum(Str(NumC), 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
    
    DC "3", "20", "1", pdFec  ' ADM/TRAB
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    stRegPPMN = stRegPPMN + CLng(NumC)
    stPPMN = stPPMN + CCur(SalC)
    sCad = sCad & Space(8) & "ADM/Trabajador      " & Space(5) & JDNum(Str(NumC), 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "3", "20", "2", pdFec  ' ADM/TRAB
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    stRegPPME = stRegPPME + CLng(NumC)
    stPPME = stPPME + CCur(SalC)
    sCad = sCad & Space(8) & "ADM/Trabajador      " & Space(5) & JDNum(Str(NumC), 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
    
    'Subtotal PP
    sCad = sCad & Space(15) & NegritaOn & "Sub Total" & Space(9) & JDNum(Str(stRegPPMN), 6, False, 6, 0) & Space(7) & JDNum(Str(stPPMN), 13, True, 11, 2)
    sCad = sCad & Space(15) & "Sub Total" & Space(8) & JDNum(Str(stRegPPME), 10, False, 10, 0) & Space(8) & JDNum(Str(stPPME), 13, True, 11, 2) + NegritaOff + Chr$(10)
    
    DC "A", "01", "1", pdFec  ' AGRICOLA
    TotalCMN = TotalCMN + CCur(SalC)
    RegCMN = RegCMN + CLng(NumC)
    sCad = sCad & Space(8) & "Agricola            " & Space(5) & JDNum(Str(NumC), 6, False, 6, 0) & Space(7) & JDNum(SalC, 13, True, 11, 2)
    DC "A", "01", "2", pdFec  ' AGRICOLA
    TotalCME = TotalCME + CCur(SalC)
    RegCME = RegCME + CLng(NumC)
    sCad = sCad & Space(8) & "Agricola            " & Space(5) & JDNum(Str(NumC), 9, False, 9, 0) & Space(8) & JDNum(SalC, 13, True, 11, 2) & Chr$(10)
    
    ' Total Créditos
    sCad = sCad & NegritaOn & Space(15) & "Total    " & Space(9) & JDNum(Str(RegCMN), 6, False, 6, 0) & Space(7) & JDNum(Str(TotalCMN), 13, True, 11, 2)
    sCad = sCad & Space(15) & "Total    " & Space(8) & JDNum(Str(RegCME), 10, False, 10, 0) & Space(8) & JDNum(Str(TotalCME), 13, True, 11, 2) + NegritaOff + Chr$(10)
    
    'Devuelve la cadena generada
    Creditos = sCad
End Function




Private Function DP(ByVal Estado As String, Optional ByVal dFec As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim bHoy As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    oCon.AbreConexion
    
    bHoy = True
  '  If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True
    
    If bHoy Then
        If gsCodAge <> "" Then
            sqlAdd = " And P.cCtaCod Like '___" & Right(gsCodAge, 2) & "%'"
        Else
            sqlAdd = ""
        End If
    Else
        If gsCodAge <> "" Then
            sqlAdd = " And cCodAge Like '%" & Right(gsCodAge, 2) & "'"
        Else
            sqlAdd = ""
        End If
    End If
    
    Select Case Estado
        Case "V"  'Desembolsado, vencido, en remate y renovado
            If bHoy Then
                If lnTipoPigno = 2 Then
                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull((Select Sum(PRD.nSaldo) nSaldo " _
                        & " From ColocPigno CO Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
                        & " Where PRD.nPRDEstado In ('2802','2803','2804') " & sqlAdd & "),0) nSaldo From" _
                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
                        & "  Where PRD.nPRDEstado In ('2802','2803','2804','2807') " & sqlAdd & " Group By CO.cCtaCod) As REP"
                Else
'                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull(Sum(nSaldo),0) nSaldo From" _
'                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
'                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
'                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
'                        & "  Where PRD.nPRDEstado In ('2101','2104','2106','2107') " & sqlAdd & " Group By CO.cCtaCod) As REP"
                        
                        SQL = " select nnum=count(*),noro=isnull(sum(oroneto),0 ),nsaldo=isnull(sum(saldo),0 ) "
                        SQL = SQL & " from (select p.cctacod,oroneto=sum(cd.npesoneto), saldo=sum(p.nsaldo) "
                        SQL = SQL & "    from colocpigjoyadet cd "
                        SQL = SQL & "     join producto p on p.cctacod=cd.cctacod "
                        SQL = SQL & "    where p.nprdestado In ('2101','2104','2106','2107')" & sqlAdd
                        SQL = SQL & "  group by p.cctacod     )  as rep "

                        
                        
                End If
            Else
                SQL = "Select Sum(nNumCredVig) nNum, Sum(nOroVig) nOro, Sum(nCapVig) nSaldo FROM ColocEstadDiaPrenda " _
                    & "WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
            End If
           
        Case "A" ' Adjudicado
            If bHoy Then
                If lnTipoPigno = 2 Then
                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull((Select Sum(PRD.nSaldo) nSaldo " _
                        & " From ColocPigno CO Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
                        & " Where PRD.nPRDEstado In ('2812') " & sqlAdd & "),0) nSaldo From" _
                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
                        & "  Where PRD.nPRDEstado In ('2812') " & sqlAdd & " Group By CO.cCtaCod) As REP"

                Else
'                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull(Sum(nSaldo),0) nSaldo From" _
'                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
'                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
'                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
'                        & "  Where PRD.nPRDEstado In ('2108') " & sqlAdd & " Group By CO.cCtaCod) As REP"

                      SQL = " select nnum=count(*),noro=isnull(sum(oroneto),0 ),nsaldo=isnull(sum(saldo),0 ) "
                    SQL = SQL & " from (select p.cctacod,oroneto=sum(cd.npesoneto), saldo=sum(p.nsaldo) "
                    SQL = SQL & "    from colocpigjoyadet cd "
                    SQL = SQL & "     join producto p on p.cctacod=cd.cctacod "
                    SQL = SQL & "    where p.nprdestado='2108' " & sqlAdd
                    SQL = SQL & "  group by p.cctacod     )  as rep "
   


                End If
            Else
                SQL = "Select Sum(nNumCredAdj) nNum, Sum(nOroAdj) nOro, Sum(nCapAdj) nSaldo FROM ColocEstadDiaPrenda " _
                    & "WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
            End If
        Case "D" 'Diferido
            If bHoy Then
                SQL = "SELECT IsNull(count(cCodCta),0) nNum, IsNull(sum(nOroNeto),0) nOro, IsNull(sum(nSaldoCap),0) nSaldo " _
                    & "FROM CREDPRENDA WHERE cEstado IN ('2')"
                If lnTipoPigno = 2 Then
'                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull((Select Sum(PRD.nSaldo) nSaldo from ColocPigno CO Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod Where PRD.nPRDEstado In ('2102')),0) nSaldo From" _
'                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
'                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
'                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
'                        & "  Where PRD.nPRDEstado In ('2102') " & sqlAdd & " Group By CO.cCtaCod) As REP"
'*****
             


                Else
'                    SQL = " Select Count(*) nNum, IsNull(Sum(nOro),0) nOro, IsNull(Sum(nSaldo),0) nSaldo From" _
'                        & " (Select CO.cCtaCod, Sum(nPesoNeto) nOro, Sum(PRD.nSaldo) nSaldo from ColocPigno CO" _
'                        & " Inner Join Producto PRD On CO.cCtaCod = PRD.cCtaCod " _
'                        & " Inner Join ColocPigJoyaTasacion CPT On CO.cCtaCod = CPT.cCtaCod And CO.nTipoTasacion = CPT.nTipoTasacion" _
'                        & "  Where PRD.nPRDEstado In ('2108') " & sqlAdd & " Group By CO.cCtaCod) As REP"

                    SQL = " select nnum=count(*), noro=isnull(sum(oroneto),0 ),nsaldo=isnull(sum(saldo),0 ) "
                    SQL = SQL & " from (select p.cctacod,oroneto=sum(cd.npesoneto), saldo=sum(p.nsaldo) "
                    SQL = SQL & "    from colocpigjoyadet cd "
                    SQL = SQL & "     join producto p on p.cctacod=cd.cctacod "
                    SQL = SQL & "    where p.nprdestado='2102' " & sqlAdd
                    SQL = SQL & "  group by p.cctacod     )  as rep "

             

                End If
            Else
                SQL = "Select Sum(nNumCredDif) nNum, Sum(nOroDif) nOro, nSaldo = 0 FROM ColocEstadDiaPrenda " _
                    & "WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
            End If
    End Select
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
        DP = "0"
        NumP = "0"
        SalP = "0"
    Else
        DP = Str(IIf(IsNull(rs!nNum), 0, rs!nNum))
        NumP = Str(IIf(IsNull(rs!nOro), 0, rs!nOro))
        SalP = Str(IIf(IsNull(rs!nSaldo), 0, rs!nSaldo))
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function BL(ByVal Moneda As String, ByVal ValMon As String, ByVal TipMon As String, ByVal dFec As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    SQL = " Select IsNull(Sum(IsNull(MUE.nMonto,0)),0) Monto from MovUserEfectivo MUE" _
        & " Inner Join Mov M On MUE.nMovNro = M.nMovNro Inner Join Efectivo E On E.cEfectivoCod = MUE.cEfectivoCod " _
        & " Where M.cOpeCod in ('901007', '901016') And Left(E.cEfectivoCod ,2) = '" & Moneda & TipMon & "' And nEfectivoValor = " & ValMon & " And Left(M.cMovNro,8) = '" & Format(dFec, gsFormatoMovFecha) & "'"
    
    If gsCodAge <> "" Then
        SQL = SQL & " And Substring(M.cMovNro, 18,2) = '" & Right(gsCodAge, 2) & "'"
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
      BL = "0"
      Select Case Moneda
           Case gMonedaNacional
             CA = 0
           Case gMonedaExtranjera
             CAE = 0
       End Select
      
    Else
      Select Case Moneda
         Case gMonedaNacional:
            CA = rs!Monto / Val(ValMon)
            BL = CCur(rs!Monto)
            TB = TB + BL
         Case gMonedaExtranjera:
            If Val(ValMon) = 0 Then ValMon = 1
            CAE = rs!Monto / Val(ValMon)
            BL = CCur(rs!Monto)
            TBD = TBD + BL
      End Select
    End If
    rs.Close
    Set rs = Nothing
End Function


Private Sub DC(ByVal TipCre As String, ByVal Fondo As String, ByVal pnMoneda As Moneda, ByVal dFec As Date)
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim bHoy As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
     
    oCon.AbreConexion
    
    bHoy = False
    If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True
    'substring(cLineaCred,3,1) = '1' AND
    Select Case TipCre
        Case "2", "2','4"  ' CPE
            If bHoy Then
                SQL = "SELECT IsNull(COUNT(PRD.cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM Colocaciones CO Inner Join Producto PRD ON CO.cCtaCod = PRD.cCtaCod " _
                    & "WHERE substring(cLineaCred,7,1) IN ('" & TipCre & "','1') AND " _
                    & "" _
                    & "substring(cLineaCred,1,2) IN ('" & Fondo & "') AND " _
                    & "SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & "AND PRD.nPRDEstado Like '20[23]%'  And Substring(PRD.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "'"
                If gsCodAge <> "" Then
                    SQL = SQL & " And Substring(PRD.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                End If
            Else
                SQL = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM ColocEstadDiaCred " _
                    & "WHERE substring(cLineaCred,7,1) IN ('" & TipCre & "','1') AND " _
                    & " " _
                    & "substring(cLineaCred,1,2) IN ('" & Fondo & "') AND " _
                    & "SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & "AND DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
                If gsCodAge <> "" Then
                    SQL = SQL & " And cCodAge = '" & Right(gsCodAge, 2) & "'"
                End If
            End If
        Case "3" 'PP
            If bHoy Then
                SQL = " SELECT IsNull(count(PRD.cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM Colocaciones CO Inner Join Producto PRD ON CO.cCtaCod = PRD.cCtaCod " _
                    & " WHERE substring(cLineaCred,7,1) = '" & TipCre & "' AND " _
                    & " substring(cLineaCred,1,2) = '" & Fondo & "' AND " _
                    & " SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & " AND PRD.nPRDEstado Like '20[23]%'  And Substring(PRD.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "'"
                If gsCodAge <> "" Then
                    SQL = SQL & " And Substring(PRD.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                End If
            Else
                SQL = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM ColocEstadDiaCred " _
                    & "WHERE substring(cLineaCred,7,1) = '" & TipCre & "' AND " _
                    & "substring(cLineaCred,1,2) = '" & Fondo & "' AND " _
                    & "SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & "AND nSaldoCap <> 0 AND DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
                    'se agrego el nsaldocap<>0 para validacion de la misma WARO 08/06/2004
                If gsCodAge <> "" Then
                    SQL = SQL & " And cCodAge = '" & Right(gsCodAge, 2) & "'"
                End If
            End If
        Case "A" ' AGRIC
            If bHoy Then
                SQL = " SELECT IsNull(count(PRD.cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM Colocaciones CO Inner Join Producto PRD ON CO.cCtaCod = PRD.cCtaCod " _
                    & " WHERE substring(cLineaCred,1,1) IN ('1','2') AND " _
                    & " substring(cLineaCred,3,1) = '2' AND " _
                    & " SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & " AND PRD.nPRDEstado Like '20[23]%' And Substring(PRD.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "'"
                If gsCodAge <> "" Then
                    SQL = SQL & " And Substring(PRD.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                End If
                
            Else
                SQL = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM ColocEstadDiaCred " _
                    & "WHERE substring(cLineaCred,1,1) IN ('1','2') AND " _
                    & "substring(cLineaCred,3,1) = '2' AND " _
                    & "SUBSTRING(cLineaCred,5,1) = '" & pnMoneda & "' " _
                    & "AND DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
                If gsCodAge <> "" Then
                    SQL = SQL & " And cCodAge = '" & Right(gsCodAge, 2) & "'"
                End If
            End If
    End Select
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
        NumC = "0"
        SalC = "0"
    Else
        NumC = Str(rs!nNum)
        SalC = Str(rs!nSaldo)
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub DA(ByVal Producto As String, ByVal Moneda As String, ByVal dFec As Date, Optional chqProd As String = "1")
    Dim SQL As String
    Dim bHoy As Boolean
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    bHoy = False
    If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True
    
    Select Case Producto
        Case gCapAhorros
            If bHoy Then
                SQL = " Select IsNull(COUNT(cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo" _
                    & " From Producto where Substring(cCtaCod,6,4) = '" & gCapAhorros & Moneda & "' And nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')"
            Else
                'Sql = " Select nNumCtas nNum, nSaldo nSaldo" _    consulta original 3 linea WARO 29/05/2004
                '    & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                '    & " And nMoneda = '" & Moneda & "' And nProducto = '" & gCapAhorros & "'"
                                
                SQL = " Select IsNull(SUM(nNumCtas),0) nNum, IsNull(SUM(nSaldo),0) nSaldo" _
                    & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                    & " And nMoneda = '" & Moneda & "' And nProducto = '" & gCapAhorros & "'"
                'modificado las 3 lineas superiores para efecto de cuentade num cuentas y suma de ahorros
            End If
            If gsCodAge <> "" Then
                'sql = sql & " And cCodAge = '" & Right(gsCodAge, 2) & "'" cons original al 27/052004 'solo habia esta linea
                If bHoy Then                                                    ' waro agregado
                    SQL = SQL & " And Substring(Producto.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'" 'waro agregado
                Else                                                            'waro agregado
                    SQL = SQL & " And cCodAge = '" & Right(gsCodAge, 2) & "'"   'waro agregado
                End If                                                          'waro agregado
            End If
       Case gCapPlazoFijo
            If bHoy Then
                SQL = " Select IsNull(COUNT(cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo" _
                    & " From Producto where Substring(cCtaCod,6,4) In ('" & gCapPlazoFijo & Moneda & "','" & gCapCTS & Moneda & "') And nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')"
            Else
               ' Sql = " Select nNumCtas nNum, nSaldo nSaldo " _  consulta original al 29/05/2004
               '     & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
               '     & " And nMoneda = '" & Moneda & "'  And nProducto In ('" & gCapPlazoFijo & "','" & gCapCTS & "') "
                    
                SQL = " Select IsNull(SUM(nNumCtas),0) nNum, IsNull(SUM(nSaldo),0) nSaldo " _
                    & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                    & " And nMoneda = '" & Moneda & "'  And nProducto In ('" & gCapPlazoFijo & "','" & gCapCTS & "') "
                'modificado el 29/05/2004 por la opcion de cuenta cuentas y suma cuentas
            End If
            If gsCodAge <> "" Then
                'sql = sql & " And cCodAge = '" & Right(gsCodAge, 2) & "'" 'consulta original 27052004 solohabia esta linea
                If bHoy Then                                                    ' waro agregado
                    SQL = SQL & " And Substring(Producto.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'" 'waro agregado
                Else                                                            'waro agregado
                    SQL = SQL & " And cCodAge = '" & Right(gsCodAge, 2) & "'"   'waro agregado
                End If                                                          'waro agregado
                
            End If
       Case "Cheques"
             Select Case chqProd
                 Case gCapAhorros
                    If bHoy Then
                        SQL = " Select IsNull(COUNT(DRC.cCtaCod),0) nNum, IsNull(sum(DRC.nMonto),0) nSaldo  From DocRec DR" _
                            & " Inner Join DocRecCapta DRC On" _
                            & " DR.nTpoDoc = DRC.nTpoDoc And DR.cNroDoc = DRC.cNroDoc And DR.cPersCod = DRC.cPersCod And DR.cIFTpo = DRC.cIFTpo" _
                            & " Where DR.nEstado = 1 And substring(DRC.cCtaCod,6,4) = '" & gCapAhorros & Moneda & "'"
                        If gsCodAge <> "" Then
                            SQL = SQL & "  And  substring(DRC.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                        End If
                    Else
                       'sql = " Select IsNull(COUNT(nNumChqVal),0) nNum, IsNull(sum(nMonChqVal),0) nSaldo " _
                       '   & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                       '   & " And nMoneda = '" & Moneda & "' And nProducto = '" & gCapAhorros & "'"
                       'Se agrego en la sgte linea de sql la variable para la agencia
                       
                        SQL = " Select IsNull(SUM(nNumChqVal),0) nNum, IsNull(sum(nMonChqVal),0) nSaldo " _
                           & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                           & " And nMoneda = '" & Moneda & "' And nNumChqVal <> 0 " & " And nProducto = '" & gCapAhorros & "'"
                            'modificado laS 3 Lineas superiores para que no sea tomado los valores 0  WARO
                      
                         If gsCodAge = "" Then
                            SQL = SQL & "  And  cCodAge = '" & Right(gsCodAge, 2) & "'"
                        End If
                            
                    End If
                    
                 Case gCapPlazoFijo
                    If bHoy Then
                        SQL = " Select IsNull(COUNT(DRC.cCtaCod),0) nNum, IsNull(sum(DRC.nMonto),0) nSaldo  From DocRec DR" _
                            & " Inner Join DocRecCapta DRC On" _
                            & " DR.nTpoDoc = DRC.nTpoDoc And DR.cNroDoc = DRC.cNroDoc And DR.cPersCod = DRC.cPersCod And DR.cIFTpo = DRC.cIFTpo" _
                            & " Where DR.nEstado = 1 And substring(DRC.cCtaCod,6,4) = '" & gCapPlazoFijo & Moneda & "'"
                        
                        If gsCodAge <> "" Then
                            SQL = SQL & "  And  substring(DRC.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                        End If
                    
                    Else
                        SQL = " Select IsNull(Sum(nNumChqVal),0) nNum, IsNull(sum(nMonChqVal),0) nSaldo " _
                            & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                            & " And nMoneda = " & Moneda & " and nNumChqVal <>0" & " And nProducto = '" & gCapPlazoFijo & "'"
                        'modificado por WARO 29/05/2004
                        
                        If gsCodAge = "" Then
                            SQL = SQL & "  And  cCodAge = '" & Right(gsCodAge, 2) & "'"
                        End If
                    End If
                 Case gCapCTS
                    If bHoy Then
                        SQL = " Select IsNull(COUNT(DRC.cCtaCod),0) nNum, IsNull(sum(DRC.nMonto),0) nSaldo  From DocRec DR" _
                            & " Inner Join DocRecCapta DRC On" _
                            & " DR.nTpoDoc = DRC.nTpoDoc And DR.cNroDoc = DRC.cNroDoc And DR.cPersCod = DRC.cPersCod And DR.cIFTpo = DRC.cIFTpo" _
                            & " Where DR.nEstado = 1 And substring(DRC.cCtaCod,6,4) = '" & gCapCTS & Moneda & "'"
                        
                        If gsCodAge <> "" Then
                            SQL = SQL & "  And  substring(DRC.cCtaCod,4,2) = '" & Right(gsCodAge, 2) & "'"
                        End If
                    Else
                        SQL = " Select IsNull(sum(nNumChqVal),0) nNum, IsNull(sum(nMonChqVal),0) nSaldo " _
                            & " FROM CapEstadSaldo WHERE DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                            & " And nMoneda = '" & Moneda & "' And nProducto = '" & gCapCTS & "'"


                        If gsCodAge = "" Then
                            SQL = SQL & "  And  cCodAge = '" & Right(gsCodAge, 2) & "'"
                        End If
                    End If
              End Select
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = oCon.CargaRecordSet(SQL)
    Set rs.ActiveConnection = Nothing
    If rs.EOF And rs.BOF Then
        NumA = "0"
        SalA = "0"
    Else
        NumA = IIf(IsNull(rs!nNum), "0", Str(rs!nNum))
        SalA = IIf(IsNull(rs!nSaldo), "0", Str(rs!nSaldo))
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Function SaldosInstFinanc(ByVal Moneda As Moneda, ByVal dFec As Date, _
            ByVal sProducto As String, ByVal sInstFinanc As PersPersoneria) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim bHoy As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim RegCmacS As Integer
    
    oCon.AbreConexion
    
    bHoy = False
    If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True
    
    If bHoy Then
        SQL = " Select IsNull(Count(PRD.cCtaCod),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM Producto PRD" _
            & " Inner Join Captaciones CAP ON PRD.cCtaCod = CAP.cCtaCod Where " _
            & " nPersoneria = '" & sInstFinanc & "' And PRD.nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')" _
            & " And Substring(PRD.cCtaCod, 6,4) = '" & Trim(sProducto) & Trim(Str(Moneda)) & "'"
    Else
        SQL = "Select nNumCtas nNum, nSaldo nSaldo From CapEstadSaldo  " _
            & "Where DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
            & "And nMoneda = " & Moneda & " And nProducto = " & sProducto & " And nPersoneria = " & sInstFinanc
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = oCon.CargaRecordSet(SQL)
    Set rs.ActiveConnection = Nothing
    If rs.EOF And rs.BOF Then
        RegCmacS = 0
        SaldosInstFinanc = "0"
    Else
        RegCmacS = IIf(IsNull(rs!nNum), 0, rs!nNum)
        RegCmacStemp = RegCmacS
        SaldosInstFinanc = IIf(IsNull(rs!nSaldo), "0", Str(rs!nSaldo))
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function GetRepCapNMejoresClientes(pnNumero As Long, lnTipoCambio As Double, psAgeCod As String, pgsEmpresa As String, pgsNomAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lnContador As Long
    
    Dim lsCadena As String
    
    Dim lsCod As String * 10
    Dim lsNom As String * 35
    Dim lsDir As String * 35
    Dim lsSaldo As String * 20
    Dim lsFono As String * 10
    Dim lsNac As String * 10
    Dim lsZon As String * 25
    Dim lsContador As String * 4
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim sqlAdd As String
    Dim lsAgenciasComentario As String
    lsCadena = ""
    
    'GetTipCambio pdIni
     
    'lnTipoCambio = gnTipCambio
     
 '  sql = " Select TOP " & pnNumero & " TA.cCodPers, P.cNomPers, TA.nSaldo, P.cDirPers, P.cTelPers," _
        & " P.dFecNac, ISNULL(Z.cDesZon,'') Zona FROM (" _
        & "     Select T.cCodPers, SUM(T.nSaldo) nSaldo FROM (" _
        & "         Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1)" _
        & "             WHEN '1' THEN A.nSaldCntAC WHEN '2' THEN A.nSaldCntAC*" & lnTipoCambio & " END" _
        & "         FROM AhorroCConsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta WHERE A.cEstCtaAC" _
        & "         NOT IN ('C','U') AND PC.cRelaCta = 'TI'" _
        & " Union" _
        & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1)" _
        & "         WHEN '1' THEN A.nSaldCntPF WHEN '2' THEN A.nSaldCntPF*" & lnTipoCambio & " END" _
        & "         from PlazoFijoConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where" _
        & "         A.cEstCtaPF not in ('C','U') and PC.cRelaCta = 'TI'" _
        & " Union" _
        & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1)" _
        & "         WHEN '1' THEN A.nSaldCntCTS WHEN '2' THEN A.nSaldCntCTS*" & lnTipoCambio & " END" _
        & "         from CTSConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where" _
        & "         A.cEstCtaCTS not in ('C','U') and PC.cRelaCta = 'TI'" _
        & "     ) T GROUP BY T.cCodPers" _
        & " ) TA INNER JOIN " & gcCentralPers & "Persona P LEFT JOIN " & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon" _
        & " ON TA.cCodPers = P.cCodPers ORDER BY TA.nSaldo DESC"
    
'    oCon.AbreConexionRemota "07", , , "05"
    
  If psAgeCod = "" Then
    sqlAdd = ""
    lsAgenciasComentario = ""
  Else
    sqlAdd = " And SubString(A.cCtaCod,4,2) = '" & psAgeCod & "'"
    lsAgenciasComentario = " Agencia = " & psAgeCod
  End If
  
  SQL = " Select TOP " & pnNumero & " TA.cPersCod cCodPers, P.cPersNombre cNomPers, TA.nSaldo, P.cPersDireccDomicilio cDirPers, P.cPersTelefono cTelPers," _
        & " P.dPersNacCreac dFecNac, ISNULL(Z.cUbiGeoDescripcion,'') Zona FROM (" _
        & "     Select T.cPersCod , SUM(T.nSaldo) nSaldo FROM (" _
        & "         Select PC.cPersCod , A.cCtaCod, nSaldo = CASE SUBSTRING(A.cCtaCod,9,1)" _
        & "             WHEN '1' THEN A.nSaldo WHEN '2' THEN A.nSaldo*" & lnTipoCambio & " END" _
        & "         FROM Producto A INNER JOIN ProductoPersona PC ON A.cCtaCod = PC.cCtaCod WHERE A.nPrdEstado" _
        & "         NOT IN ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "') AND PC.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & "" _
        & "         And Substring(A.cCtaCod,6,3) In ('" & Producto.gCapAhorros & "','" & Producto.gCapPlazoFijo & "','" & Producto.gCapCTS & "') " & sqlAdd & ") T GROUP BY T.cPersCod" _
        & " ) TA INNER JOIN Persona P" _
        & " LEFT JOIN UbicacionGeografica Z ON P.cPersDireccUbiGeo = Z.cUbiGeoCod" _
        & " ON TA.cPersCod = P.cPersCod ORDER BY TA.nSaldo DESC"
    
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(SQL)
    
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("MEJORES " & pnNumero & " MEJORES CLIENTES" & lsAgenciasComentario, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, , False)
    lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;7; ;5;Nombre;15; ;22;Dirección;15; ;22;Saldo;18; ;4;Fono;5; ;7;Fec.Nac;8; ;4;Zona;15; ;12;", lnItem)
    
    lnContador = 0
    
    While Not rs.EOF
        lnContador = lnContador + 1
        lsCod = rs!cCodPers
        lsNom = rs!cNomPers
        lsDir = rs!cDirPers
        RSet lsSaldo = Format(rs!nSaldo, "#,##0.00")
        lsFono = rs!cTelPers & ""
        lsNac = Format(rs!dFecNac, gsFormatoFechaView)
        lsZon = rs!Zona
        
        lsCadena = lsCadena & Format(lnContador, "0000") & "  " & lsCod & "  " & lsNom & "  " & lsDir & "  " & lsSaldo & "  " & lsFono & "  " & lsNac & "  " & lsZon & Chr(10)
        
        lnItem = lnItem + 1
        
        If lnItem > 54 Then
            lnItem = 0
            lsCadena = lsCadena & Chr(12)
            lsCadena = lsCadena & CabeceraPagina("MEJORES " & pnNumero & " MEJORES CLIENTES" & lsAgenciasComentario, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, , False)
            lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;7; ;5;Nombre;15; ;22;Dirección;15; ;22;Saldo;18; ;4;Fono;5; ;7;Fec.Nac;8; ;4;Zona;15; ;12;", lnItem)
        End If
        
        rs.MoveNext
    Wend
    
    rs.Close
    
    GetRepCapNMejoresClientes = lsCadena
End Function

Public Function CuentasAperturadas(pAge As String, pTitulo As String, pTipoCta As String, pOrden As Integer, pMoneda As String, dFecIni As Date, dFecFin As Date, pnMontoIni As Currency, pnMontoFin As Currency) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim lnTotal As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsCtaCod As String * 18
    Dim lsSaldo As String * 15
    Dim lsApertura As String * 10
    Dim lsVencimiento As String
    Dim lsPlazo As String * 5
    Dim lsNombre As String * 40
    Dim lsDireccion As String * 40
    Dim lsFono As String * 6
    Dim sqlAdd As String
    Dim lsCtaCodAnt As String
    
    If pAge <> "" Then
        sqlAdd = " and substring(Cp.cCtaCod,4,2) = '" & pAge & "'"
    Else
        sqlAdd = ""
    End If
    
    
    oCon.AbreConexion
    Select Case pTipoCta
    Case "232"
    SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacAhorros  CA on  CA.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where substring(Cp.cCtaCod,6,3)='" & pTipoCta & "' and PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1000,1100,1101,1200,1201) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & " " & sqlAdd & " and CA.bOrdPag=" & pOrden
    SQL = SQL & " Order by Cp.cCtacod "
    
    Case "233"
    SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacPlazoFijo  CPF on  CPF.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1000,1100,1101,1200,1201) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & "  " & sqlAdd & " "
    SQL = SQL & " Order by Cp.cCtacod "
    
    Case "234"
        SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacCTS CTS on  CTS.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1000,1100,1101,1200,1201) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & "  " & sqlAdd & "  "
    SQL = SQL & " Order by Cp.cCtacod "
    End Select
    
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & Format(gdFecSis, "dddd, dd mmmm yyyy") & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
     lsCad = lsCad & Space(20) & pTitulo & "EN " & IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
     lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & "MONTO" & Space(5) & "F. APER" & Chr(10)
     lsCad = lsCad & String(100, "=") & Chr(10)
     Do While Not rs.EOF
        lsCtaCod = rs!Cctacod
        
        If lsCtaCodAnt = rs!Cctacod Then
            lsSaldo = ""
            lsVencimiento = ""
        Else
            RSet lsSaldo = Format(rs!nSaldoDisp, "#,##0.00")
            lsVencimiento = Format(rs!dApertura, gsFormatoFechaView)
        End If
        lsNombre = rs!cPersNombre
                
        lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsSaldo & Space(2) & lsVencimiento & Chr(10)
        
        lsCtaCodAnt = rs!Cctacod
        rs.MoveNext
        vLineas = vLineas + 1
        
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
            lsCad = lsCad & Space(20) & pTitulo & "EN " & IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & Space(40) & "MONTO" & Space(5) & "F. APER" & Chr(10)
            lsCad = lsCad & String(160, "=") & Chr(10)
            vLineas = 1
        End If
     Loop
    
    End If
    rs.Close
    Set rs = Nothing
    CuentasAperturadas = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Public Function CuentasCanceladas(pAge As String, pTitulo As String, pTipoCta As String, pOrden As Integer, pMoneda As String, dFecIni As Date, dFecFin As Date, pnMontoIni As Currency, pnMontoFin As Currency) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim lnTotal As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsCtaCod As String * 18
    Dim lsSaldo As String * 15
    Dim lsApertura As String * 10
    Dim lsVencimiento As String
    Dim lsPlazo As String * 5
    Dim lsNombre As String * 40
    Dim lsDireccion As String * 40
    Dim lsFono As String * 6
    Dim sqlAdd As String
    
    If pAge <> "" Then
        sqlAdd = " and substring(Cp.cCtaCod,4,2) = '" & pAge & "'"
    Else
        sqlAdd = ""
    End If
    
    oCon.AbreConexion
    Select Case pTipoCta
    Case "232"
    SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacAhorros  CA on  CA.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where substring(Cp.cCtaCod,6,3)='" & pTipoCta & "' and PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1300,1301,1400,1401) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & sqlAdd & " and CA.bOrdPag=" & pOrden
    SQL = SQL & " Order by Cp.cCtacod "
    
    Case "233"
    SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacPlazoFijo  CPF on  CPF.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1300,1301,1400,1401) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & sqlAdd
    SQL = SQL & " Order by Cp.cCtacod "
    
    Case "234"
        SQL = "Select Cp.cCtaCod, P.cPersNombre, Cp.nSaldoDisp, "
    SQL = SQL & " Cp.dApertura from Captaciones Cp "
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join CaptacPCTS  CTS on  CTS.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join ProductoPersona PP on PP.cCtaCod=Cp.cCtaCod "
    SQL = SQL & " Inner Join Persona P on  P.cPersCod=PP.cPersCod "
    SQL = SQL & " where PP.nPrdPersRelac=" & gCapRelPersTitular
    SQL = SQL & " and substring(Cp.cCtaCod,9,1)='" & pMoneda & "' and Pro.nPrdEstado in (1300,1301,1400,1401) and "
    SQL = SQL & " Cp.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' and '" & Format(dFecFin, gsFormatoFecha) & "' "
    SQL = SQL & sqlAdd
    SQL = SQL & " Order by Cp.cCtacod "
    End Select
    
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad & Chr(10)
     lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & Format(gdFecSis, "dddd, dd mmmm yyyy") & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
     lsCad = lsCad & Space(20) & pTitulo & "EN " & IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
     lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & "MONTO" & Space(5) & "F. APER" & Chr(10)
     lsCad = lsCad & String(100, "=") & Chr(10)
     Do While Not rs.EOF
        lsCtaCod = rs!Cctacod
        RSet lsSaldo = Format(rs!nSaldoDisp, "#,##0.00")
        lsVencimiento = Format(rs!dApertura, gsFormatoFechaView)
        lsNombre = rs!cPersNombre
                
        lsCad = lsCad & lsCtaCod & Space(2) & lsNombre & Space(2) & lsSaldo & Space(2) & Format(rs!dApertura, gsFormatoFechaView) & Chr(10)
        
        rs.MoveNext
        vLineas = vLineas + 1
        
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & gsNomAge & Space(5) & "Sección Ahorro" & Space(5) & "Trujillo " & ArmaFecha(gdFecSis) & Space(4) & "Pagina: " & Format(vPag, "####") & Chr(10) & Chr(10)
            lsCad = lsCad & Space(20) & pTitulo & "EN " & IIf(pMoneda = Moneda.gMonedaNacional, "SOLES", "DOLARES") + Chr(10) + Chr(10)
            lsCad = lsCad & "CUENTA " & Space(25) & "NOMBRE" & Space(35) & Space(40) & "MONTO" & Space(5) & "F. APER" & Chr(10)
            lsCad = lsCad & String(160, "=") & Chr(10)
            vLineas = 1
        End If
     Loop
    
    End If
    rs.Close
    Set rs = Nothing
    CuentasCanceladas = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Public Function ReportesExtornos(pAge As String, pTitulo As String, pMoneda As String, _
        dFecIni As Date, dFecFin As Date, pnMontoIni As Currency, pnMontoFin As Currency, _
        sNomAge As String) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql1 As String
    Dim CuentaA As String * 18
    Dim vLineas As Integer
    Dim vPag As Integer
    Dim lnTotal As Currency
    Dim RCapital As Currency
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsAgencias As String * 19
    Dim lsCtaCod As String * 18
    Dim lsMovNro As String * 10
    Dim lsMonto As String * 10
    
    Dim lsNomOpe As String, lsDoc As String, lsUsuario As String
    Dim sTitulo2 As String
    Dim lsFecha As Date
    
    Dim lnMonedaAnt As Integer
    
    Dim lnCorr As Integer
    
    If dFecIni = dFecFin Then
        sTitulo2 = "DEL " & Format$(dFecIni, gcFormatoFechaView)
    Else
        sTitulo2 = "DEL " & Format$(dFecIni, gcFormatoFechaView) & " AL " & Format$(dFecFin, gcFormatoFechaView)
    End If
    
    oCon.AbreConexion
    
    SQL = "Select M.nMovNro,  O.cOpeDesc + ' ' + Replace(Replace(m.cMovDesc,Char(10),' '),Char(13),' ') cOpeDesc, IsNull((Select cDocDesc from Documento where nDocTpo=MD.nDocTpo),IsNull((Select cUsuDest from MovHabDevCajero MHC Where MHC.nMovNro = M.nMovNro),'')) Documento , M.cMovNro, substring(M.cMovNro,22,4) Usuario,"
    SQL = SQL & " IsNull(MC.cCtaCod,(Select Top 1 cNumRecibo FROM MovServiciosDet MM Where MM.nMovNro = M.nMovNro)) CTA, Agencia=(Select cAgeDescripcion from Agencias where substring(M.cMovNro,18,2)= cAgeCod), "
    SQL = SQL & "     IsNull((Select nMovImporte From MovHabDevCajero A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "      IsNull((Select nMonto From MovCap A Where M.nMovNro = A.nMovNro And A.cCtaCod = MC.cCtaCod),"
    SQL = SQL & "       IsNull((Select nMonto from MovCol A Where M.nMovNro = A.nMovNro And A.cCtaCod = MCO.cCtaCod),"
    SQL = SQL & "        IsNull((Select nMonto from MovServicios A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "         IsNull((Select nMovImporte from MovOpeVarias A Where M.nMovNro = A.nMovNro),0))))) Monto,"
    SQL = SQL & "     IsNull((Select Top 1 nMoneda From MovHabDevCajero A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "      IsNull((Select Top 1 Substring(A.cCtaCod,9,1) From MovCap A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "       IsNull((Select Top 1 Substring(A.cCtaCod,9,1) from MovCol A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "        IsNull((Select Top 1 nMoneda from MovServicios A Where M.nMovNro = A.nMovNro),"
    SQL = SQL & "         IsNull((Select Top 1 nMoneda from MovOpeVarias A Where M.nMovNro = A.nMovNro),0))))) Moneda,"
    SQL = SQL & " convert(datetime,convert(varchar(8),cMovNro,103),103)  Fecha From Mov M"
    SQL = SQL & " Left  Join MovCap MC on MC.nMovNro=M.nMovNro"
    SQL = SQL & " Left  Join MovCol MCO on MCO.nMovNro=M.nMovNro"
    SQL = SQL & " Inner Join OpeTpo O on O.copeCod = M.cOpeCod"
    SQL = SQL & " left outer join MovDoc MD on MD.nMovNro=M.nMovNro"
    SQL = SQL & " where substring(M.cOpeCod,1,1) In ('2','1','9','3','') and M.nMovFlag <> " & MovFlag.gMovFlagVigente & ""
    SQL = SQL & " And Left(cMovNro,8) Between '" & Format(dFecIni, gsFormatoMovFecha) & "' And '" & Format(dFecFin, gsFormatoMovFecha) & "'"
    SQL = SQL & " And M.cOpeCod Not Like '1[034]9[012]%' And M.cOpeCod Not Like '129%'"
    SQL = SQL & " And M.cOpeCod Not Like '159%' And M.cOpeCod Not Like '1[25]9%'"
    SQL = SQL & " And M.cOpeCod Not Like '2[3457]%' And M.cOpeCod Not Like '3[569]%'"
    SQL = SQL & " And M.cOpeCod Not Like '90900[0-3]%' And M.cOpeCod Not Like '90900[0-6]%'"
    SQL = SQL & " And M.cOpeCod Not Like '90103[0-9]%' And M.cOpeCod Not Like '90102[1-9]%'"
    SQL = SQL & " And M.cOpeCod Not Like '90003[6-9]%' And M.cOpeCod Not Like '90004[4-6]%'"
    If pAge <> "" Then SQL = SQL & " and substring(M.cMovNro,18,2) = " & pAge
    SQL = SQL & " ORDER BY Moneda, M.cMovNro"
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 0
     
     lnCorr = 0
     
     lnMonedaAnt = -1
     
     Do While Not rs.EOF
        If lnMonedaAnt <> rs!Moneda Then
             vPag = vPag + 1
             If lsCad <> "" Then
                lsCad = lsCad & String(151, "-") & Chr(10)
                lsCad = lsCad & Chr(12)
             End If
             lsCad = lsCad & Chr(10)
             lsCad = lsCad & lsEmpresa & IIf(rs!Moneda = 1, "-SOLES  ", IIf(rs!Moneda = 2, "-DOLARES", "        ")) & Space(52) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Space(20) & "Pagina: " & Format(vPag, "####") & Chr(10)
             lsCad = lsCad & sNomAge & Chr$(10)
             lsCad = lsCad & Space(Int((151 - Len(pTitulo)) / 2)) & pTitulo & Chr(10)
             lsCad = lsCad & Space(Int((151 - Len(sTitulo2)) / 2)) & sTitulo2 & Chr(10) & Chr$(10)
        
             lsCad = lsCad & String(151, "=") & Chr(10)
             lsCad = lsCad & "ITEM    NRO.MOV  OPERACION " & Space(37) & "DOCUMENTO" & "            MONTO " & Space(4) & "USU." & Space(5) & "NRO CUENTA" & Space(10) & "AGENCIA" & Space(10) & "   FECHA" & Chr(10)
             lsCad = lsCad & String(151, "=") & Chr(10)
        End If
        
        
        lsNomOpe = rs!cOpeDesc
        lsDoc = rs!Documento
        lsUsuario = rs!Usuario
        lsFecha = Format(rs!Fecha, gsFormatoFechaView)
        CuentaA = rs!Cta & ""
        lsAgencias = rs!AGENCIA & ""
        lnCorr = lnCorr + 1
        RSet lsMovNro = rs!nMovNro
        RSet lsMonto = Format(rs!Monto, "#,##0.00")
        
        lsCad = lsCad & Format(lnCorr, "0000") & Space(2) & lsMovNro & "  " & ImpreFormat(lsNomOpe, 45, 0) & Space(2) & ImpreFormat(lsDoc, 15, 0) & Space(2) & lsMonto & Space(2) & lsUsuario & Space(2) & CuentaA & Space(2) & lsAgencias & Space(2) & lsFecha & Chr(10)
        
        lnMonedaAnt = rs!Moneda
        rs.MoveNext
        vLineas = vLineas + 1
        
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad & Chr(12)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & Chr(10)
            lsCad = lsCad & Chr(10)
            
            lsCad = lsCad & lsEmpresa & IIf(rs!Moneda = 1, "-SOLES  ", IIf(rs!Moneda = 2, "-DOLARES", "        ")) & Space(52) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Space(20) & "Pagina: " & Format(vPag, "####") & Chr(10)
            lsCad = lsCad & sNomAge & Chr$(10)
            lsCad = lsCad & Space(Int((151 - Len(pTitulo)) / 2)) & pTitulo & Chr(10)
            lsCad = lsCad & Space(Int((151 - Len(sTitulo2)) / 2)) & sTitulo2 & Chr(10) & Chr$(10)
            
            lsCad = lsCad & String(151, "=") & Chr(10)
            lsCad = lsCad & "ITEM    NRO.MOV  OPERACION " & Space(37) & "DOCUMENTO" & "            MONTO " & Space(4) & "USU." & Space(5) & "NRO CUENTA" & Space(10) & "AGENCIA" & Space(10) & "   FECHA" & Chr(10)
            lsCad = lsCad & String(151, "=") & Chr(10)
            vLineas = 1
        End If
     Loop
     lsCad = lsCad & String(151, "-") & Chr(10)
    
    
    End If
    rs.Close
    Set rs = Nothing
    ReportesExtornos = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Private Function DCaja(ByVal dFec As Date, Optional sAgencia As String = "", _
            Optional nMoneda As Moneda = gMonedaExtranjera) As String
    
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String
    
    oCon.AbreConexion
    
    If sAgencia = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & sAgencia & "'"
    End If
    
    SQL = "Select ISNULL(Sum(E.nMonto),0) Caja from MovUserEfectivo E JOIN Mov M On E.nMovNro = M.nMovNro " _
        & "Where M.cMovNro Like '" & Format$(dFec, gsFormatoMovFecha) & "%' And M.nMovFlag = 0 And " _
        & "E.cEfectivoCod Like '" & nMoneda & "%' And M.cOpeCod IN ('" & gOpeHabCajRegEfect & "','" & gOpeBoveAgeRegEfect & "') " & sqlAdd
    
    Set rs = oCon.CargaRecordSet(SQL)
    
    If rs.EOF And rs.BOF Then
        DCaja = "0"
    Else
        DCaja = Format$(rs!Caja, "#,#00.00")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function DColocaciones(ByVal dFec As Date, Optional sAgencia As String = "", _
        Optional nMoneda As Moneda = gMonedaExtranjera, Optional bHoy As Boolean = False) As String
    
    Dim ssql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlAdd As String, sqlAge
        
    If sAgencia = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And SUBSTRING(P.cCtaCod,4,2) = '" & sAgencia & "'"
        If Not bHoy Then sqlAdd = " And cCodAge = '" & sAgencia & "'"
    End If
    
    oCon.AbreConexion
    If bHoy Then
        ssql = "Select ISNULL(SUM(nSaldo),0) nSaldo From Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
            & "Where P.nPrdEstado IN (2020,2021,2022,2023,2030,2031,2032,2033,2802,2803,2804) And " _
            & "SUBSTRING(P.cCtaCod,9,1) = '" & nMoneda & "' " & sqlAdd
    Else
        ssql = "SELECT IsNull(SUM(nSaldoCap),0) nSaldo FROM ColocEstadDiaCred Where " _
        & "SUBSTRING(cLineaCred,6,1) = '" & nMoneda & "' And DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
    End If
    
    Set rs = oCon.CargaRecordSet(ssql)
    
    If rs.EOF And rs.BOF Then
      DColocaciones = "0.00"
    Else
      DColocaciones = Format$(rs!nSaldo, "#,#00.00")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function DPlazo(ByVal dFec As Date, Optional sAgencia As String = "", _
        Optional nMoneda As Moneda = gMonedaExtranjera, Optional bHoy As Boolean = False) As String
    
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlAdd As String
    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    If sAgencia = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cCodAge =  '" & sAgencia & "'"
        If bHoy Then sqlAdd = " And SUBSTRING(P.cCtaCod,4,2) = '" & sAgencia & "'"
    End If
            
    If bHoy Then
        SQL = "Select ISNULL(SUM(nSaldo),0) nSaldo from Producto P JOIN Captaciones C ON P.cCtaCod = C.cCtaCod " _
            & "Where P.nPrdEstado NOT IN (" & gCapEstCancelada & "," & gCapEstAnulada & ") And " _
            & "SUBSTRING(P.cCtaCod,9,1) = '" & nMoneda & "' And SUBSTRING(P.cCtaCod,6,3) " _
            & "IN ('" & gCapPlazoFijo & "','" & gCapCTS & "')  " & sqlAdd
    Else
        SQL = "Select IsNUll(Sum(nSaldo),0) nSaldo from CapEstadSaldo  " _
            & "Where nProducto In (" & gCapPlazoFijo & "," & gCapCTS & ") And nMoneda = " & nMoneda & " And " _
            & "DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
        DPlazo = "0"
    Else
        DPlazo = Format(rs!nSaldo, "#,#00.00")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function DAhorro(ByVal dFec As Date, Optional sAgencia As String = "", _
        Optional nMoneda As Moneda = gMonedaExtranjera, Optional bHoy As Boolean = False) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Dim sqlAdd As String
        
    
    If sAgencia = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cCodAge =  '" & sAgencia & "'"
        If bHoy Then sqlAdd = " And SUBSTRING(P.cCtaCod,4,2) = '" & sAgencia & "'"
    End If
            
    If bHoy Then
        SQL = "Select ISNULL(SUM(nSaldo),0) nSaldo from Producto P JOIN Captaciones C ON P.cCtaCod = C.cCtaCod " _
            & "Where P.nPrdEstado NOT IN (" & gCapEstCancelada & "," & gCapEstAnulada & ") And " _
            & "SUBSTRING(P.cCtaCod,9,1) = '" & nMoneda & "' And SUBSTRING(P.cCtaCod,6,3) = '" & gCapAhorros & "' " & sqlAdd
    Else
        SQL = "Select IsNUll(Sum(nSaldo),0) nSaldo from CapEstadSaldo  " _
            & "Where nProducto = '" & gCapAhorros & "' And nMoneda = " & gMonedaExtranjera & " And " _
            & "DateDiff(dd,dEstad,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " & sqlAdd
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
        DAhorro = "0"
    Else
        DAhorro = Format(rs!nSaldo, "#,#00.00")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function CompraVenta(TIP As String, ByVal dFec As Date) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Dim sqlAdd As String
    
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & lsAgeCod & "'"
    End If
    
    Select Case TIP
        Case "C"
            SQL = " SELECT IsNull(Sum(CV.nMovImporte),0) MontoME, IsNull(sum(CV.nMovImporte * MTC.nMovTpoCambio),0) MontoMN From Mov M Inner Join MovCompraVenta CV ON CV.nMovNro = M.nMovNro INNER JOIN  MovTpoCambio MTC ON MTC.nMovNro = M.nMovNro" _
                & " WHERE M.cOpeCod In ('" & gOpeCajeroMECompra & "','" & gOpeCajeroMECompraEsp & "') and (M.nMovFlag = " & MovFlag.gMovFlagVigente & ") " _
                & " And M.cMovNro Like '" & Format(dFec, gsFormatoMovFecha) & "%'" & sqlAdd
        Case "V"
            SQL = " SELECT IsNull(Sum(CV.nMovImporte),0) MontoME, IsNull(sum(CV.nMovImporte * MTC.nMovTpoCambio),0) MontoMN From Mov M Inner Join MovCompraVenta CV ON CV.nMovNro = M.nMovNro INNER JOIN  MovTpoCambio MTC ON MTC.nMovNro = M.nMovNro" _
                & " WHERE M.cOpeCod In ('" & gOpeCajeroMEVenta & "','" & gOpeCajeroMEVentaEsp & "') and (M.nMovFlag = " & MovFlag.gMovFlagVigente & ") " _
                & " And M.cMovNro Like '" & Format(dFec, gsFormatoMovFecha) & "%'" & sqlAdd
    End Select
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF And rs.BOF Then
       CompraVenta = "0"
       ImpSoles = "0"
    Else
       CompraVenta = Format(rs!MontoME, "#,##0.00")
       ImpSoles = Format(rs!MontoMN, "#,##0.00")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub TCambio(ByVal dFec As Date)
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    k = 1
    
    Dim sqlAdd As String
        
    If lsAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & lsAgeCod & "'"
    End If
    
    SQL = "SELECT dFecCamb, nValComp nValComp, nValVent nValVent, nValCompEsp nValCompEsp, nValVentEsp nValVentEsp FROM TipoCambio " _
        & " where DateDiff(dd,dFecCamb,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(SQL)
    
    If rs.EOF And rs.BOF Then
        DHora(1) = "00:00:00"
        TC(1) = "0"
        TV(1) = "0"
        k = 2
    Else
        Do While Not rs.EOF
            DHora(k) = Format(rs!dFecCamb, "hh:mm:ss")
            TC(k) = Format(rs!nValComp, "#,#00.00")
            TV(k) = Format(rs!nValVent, "#,#00.00")
            
            TCE(k) = Format(rs!nValCompEsp, "#,#00.00")
            TVE(k) = Format(rs!nValVentEsp, "#,#00.00")

            k = k + 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Function ImpLavDineroD(ByVal psFechaIni As String, ByVal psFechaFin As String, ByVal psNomCmac As String, _
                              ByVal psNomAge As String, ByVal psCodAge As String, pdFecSis As Date) As String
                              
Dim SQL As String
Dim oCon As DConecta
Dim rs As Recordset
Dim lsImpre As String
Dim lsFecha As String
Dim oGen As DGeneral
Dim lnMontoMenLD As Currency
Dim lsCliente As String * 30
Dim lnMontoED As Currency
Dim lnTMontoS As Currency
Dim lnTMontoED As Currency
Dim lnTMontoD As Currency
Dim lnItem As Long
Dim lsDocId As String * 11
Dim lsAgencia As String * 15
Dim lsFechaTrans As String * 19
Dim lsCtaCod As String * 18, lsFechas As String * 10
Dim lsProducto As String * 10, lsOperacion As String * 20
Dim lnTipoCambio As Currency
Dim lnContPag As Integer
Dim lnContLin As Integer
Dim lsFechaIni As String, lsFechaFin As String
Dim lnMovNro As Long

    lsFechaIni = Mid(psFechaIni, 7, 4) & Mid(psFechaIni, 4, 2) & Mid(psFechaIni, 1, 2)
    lsFechaFin = Mid(psFechaFin, 7, 4) & Mid(psFechaFin, 4, 2) & Mid(psFechaFin, 1, 2)
    
    Set oGen = New DGeneral
        lnMontoMenLD = oGen.GetParametro(2000, 2032)
    Set oGen = Nothing

    lnContPag = 1:  lnItem = 0
    lsImpre = ""
    
    Set oCon = New DConecta
    oCon.AbreConexion

    SQL = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0 ORDER BY dFecCamb "

    Set rs = oCon.CargaRecordSet(SQL)
        If Not (rs.EOF And rs.BOF) Then
            lnTipoCambio = rs!nValPond
        Else
            lnTipoCambio = 0
        End If
    Set rs = Nothing
    
    lsImpre = lsImpre & Space(5) & psNomCmac & Space(170) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "OPERACIONES EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
    lsImpre = lsImpre & Space(72) & "Pag. " & lnContPag & Chr(10)
    lsImpre = lsImpre & Space(100) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
    
    lsImpre = lsImpre & String(229, "-") & Chr(10)
    lsImpre = lsImpre & "ITEM" & Space(15) & "CLIENTE" & Space(14) & "DOC. ID" & Space(11)
    lsImpre = lsImpre & "AGENCIA" & Space(10) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO" & Space(11) & "OPERACION"
    lsImpre = lsImpre & Space(15) & "CUENTA" & Space(12) & "MONTO SOLES" & Space(5) & "EQUIV. DOLAR" & Space(5)
    lsImpre = lsImpre & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(3) & "USUARIO" & Chr(10)
    lsImpre = lsImpre & String(229, "-") & Chr(10)
    
    lnContLin = 7
        
    SQL = "SELECT nMovNro, cPersNombre, cPersIDNro,  cAgeDescripcion, Agencia, cMovNro, TipoProd, cOpeDesc, " _
        & "cCtaCod, MontoS, MontoD, Usuario FROM (" _
        & "SELECT LD.nMovNro, cPersNombre, cPersIDNro,  cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
        & "TipoProd = '', cOpeDesc, cCtaCod = '', MontoS = 0, nMovImporte MontoD, SUBSTRING(cMovNro, 22,4) Usuario " _
        & "FROM MOVCOMPRAVENTA CV INNER JOIN PERSONA P ON CV.cPersCod = P.cPersCod " _
        & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = CV.nMovNro " _
        & "INNER JOIN MOVLAVDINERO LD ON LD.nMovNro = CV.nMovNro " _
        & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
        & "WHERE cPersIDTpo in (1,2) AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "' " _
        & " AND M.nMovFlag = 0 "
        
    If psCodAge <> "" Then
        SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
    End If

    SQL = SQL & " UNION ALL " _
        & "SELECT LD.nMovNro, cPersNombre, cPersIDNro,  cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
        & "TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 232 THEN 'AHORROS' WHEN 233 THEN 'PLAZOFIJO' WHEN 234 " _
        & "THEN 'CTS' END, cOpeDesc, MC.cCtaCod, " _
        & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
        & "FROM MOVCAP MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
        & "PERSONA P ON PP.cPersCod = P.cPersCod " _
        & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
        & "INNER JOIN MOVLAVDINERO LD ON LD.nMovNro = MC.nMovNro " _
        & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
        & "WHERE cPersIDTpo IN (1,2) AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & " ' " _
        & " AND M.nMovFlag = 0 "

    If psCodAge <> "" Then
        SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
    End If

    SQL = SQL & "UNION ALL " _
        & " SELECT LD.nMovNro, cPersNombre, cPersIDNro,  cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, " _
        & " cMovNro, TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 305 THEN 'PIGNORATICIO' ELSE 'CREDITOS' END , " _
        & " cOpeDesc, MC.cCtaCod, MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & " MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & " SUBSTRING(cMovNro, 22,4) Usuario " _
        & " FROM MOVCOL MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod " _
        & " INNER JOIN PERSONA P ON PP.cPersCod = P.cPersCod INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod " _
        & " INNER JOIN MOV M ON M.nMovNro = MC.nMovNro INNER JOIN MOVLAVDINERO LD ON LD.nMovNro = MC.nMovNro " _
        & " INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
        & " WHERE cPersIDTpo IN (1,2) AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & " ' AND '" _
        & lsFechaFin & "' AND M.nMovFlag = 0 AND nPrdPersRelac = 20 "

    If psCodAge <> "" Then
        SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
    End If

    SQL = SQL & ")S ORDER BY cMovNro"
    
    Set rs = oCon.CargaRecordSet(SQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
        
            lsCliente = PstaNombre(rs!cPersNombre)
            lsDocId = rs!cPersIDNro
            lsAgencia = rs!cAgeDescripcion
            lsFechaTrans = Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4) & " " & Mid(rs!cMovNro, 9, 2) + ":" & Mid(rs!cMovNro, 11, 2) & ":" & Mid(rs!cMovNro, 13, 2)
            lsProducto = rs!TipoProd
            lsOperacion = rs!cOpeDesc
            lsCtaCod = rs!Cctacod
            If lnTipoCambio > 0 Then
                lnMontoED = rs!MontoS / lnTipoCambio
            End If

            If lnMovNro <> rs!nMovNro Then
                lnItem = lnItem + 1
                lsImpre = lsImpre & ImpreFormat(lnItem, 3, 0) & Space(3) & ImpreCarEsp(lsCliente) & Space(3) & lsDocId & Space(5)
                lsImpre = lsImpre & ImpreCarEsp(lsAgencia) & Space(5) & lsFechaTrans & Space(5) & ImpreCarEsp(lsProducto) & Space(4)
                lsImpre = lsImpre & ImpreCarEsp(lsOperacion) & Space(3) & lsCtaCod & Space(3) & ImpreFormat(rs!MontoS, 9, 2, True) & Space(5)
                lsImpre = lsImpre & ImpreFormat(lnMontoED, 9, 2, True) & Space(5) & ImpreFormat(rs!MontoD, 9, 2)
                lsImpre = lsImpre & Space(5) & ImpreFormat(lnMontoED + rs!MontoD, 9, 2) & Space(3) & rs!Usuario & Chr(10)
                lnContLin = lnContLin + 1
                lnTMontoS = lnTMontoS + rs!MontoS
                lnTMontoED = lnTMontoED + lnMontoED
                lnTMontoD = lnTMontoD + rs!MontoD
            Else
                lsImpre = lsImpre & Space(6) & ImpreCarEsp(lsCliente) & Space(3) & lsDocId & Chr(10)
                lnContLin = lnContLin + 1
            End If
            
            lnMovNro = rs!nMovNro
            
            If lnContLin >= 60 Then
                lnContPag = lnContPag + 1
                lsImpre = lsImpre & Chr(12)
                
                lsImpre = lsImpre & Space(5) & psNomCmac & Space(170) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
                lsImpre = lsImpre & Space(5) & psNomAge & Space(58) & "OPERACIONES EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
                lsImpre = lsImpre & Space(72) & "Pag. " & lnContPag & Chr(10) & Chr(10) & Chr(10)
                
                lsImpre = lsImpre & String(229, "-") & Chr(10)
                lsImpre = lsImpre & "ITEM" & Space(15) & "CLIENTE" & Space(14) & "DOC. ID" & Space(11)
                lsImpre = lsImpre & "AGENCIA" & Space(10) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO" & Space(11) & "OPERACION"
                lsImpre = lsImpre & Space(15) & "CUENTA" & Space(12) & "MONTO SOLES" & Space(5) & "EQUIV. DOLAR" & Space(5)
                lsImpre = lsImpre & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(3) & "USUARIO" & Chr(10)
                lsImpre = lsImpre & String(229, "-") & Chr(10)
                
                lnContLin = 7
                
            End If
            
            rs.MoveNext
        Loop
        
        lsImpre = lsImpre & String(229, "-") & Chr(10)
        lsImpre = lsImpre & Space(10) & "TOTAL   :  " & ImpreFormat(lnItem, 3, 0) & Space(131) & ImpreFormat(lnTMontoS, 11, 2, True) & Space(3)
        lsImpre = lsImpre & ImpreFormat(lnTMontoED, 11, 2, True) & Space(3) & ImpreFormat(lnTMontoD, 11, 2, True) & Space(3)
        lsImpre = lsImpre & ImpreFormat(lnTMontoED + lnTMontoD, 11, 2, True) & Chr(10)
        lsImpre = lsImpre & String(229, "-") & Chr(10)
        
        Set rs = Nothing
    End If

    ImpLavDineroD = lsImpre
    
End Function

Public Function ImpLavDineroM(ByVal psFechaIni As String, ByVal psFechaFin As String, ByVal psNomCmac As String, _
                              ByVal psNomAge As String, ByVal psCodAge As String, pdFecSis As Date) As String

Dim SQL As String
Dim oCon As DConecta
Dim rs As Recordset
Dim rsDet As Recordset
Dim lsImpre As String
Dim oGen As DGeneral
Dim lnMontoMenLD As Currency
Dim lsCliente As String * 40
Dim lnMontoED As Currency
Dim lnTCMontoS As Currency
Dim lnTCMontoED As Currency
Dim lnTCMontoD As Currency
Dim lnTMontoS As Currency
Dim lnTMontoED As Currency
Dim lnTMontoD As Currency
Dim lnItem As Long, lnTCItem As Long, lnTItem As Long
Dim lsDocId As String * 8
Dim lsAgencia As String * 15, lsFechaTrans As String * 19
Dim lsCtaCod As String * 18, lsFechas As String * 10
Dim lsProducto As String * 10, lsOperacion As String * 15
Dim lnTipoCambio As Currency
Dim lnContPag As Integer
Dim lnContLin As Integer
Dim lsFechaIni As String, lsFechaFin As String

    Set oGen = New DGeneral
        lnMontoMenLD = oGen.GetParametro(2000, 2033)
    Set oGen = Nothing

    lnContPag = 1:  lnItem = 0
    lnTCItem = 0: lnTItem = 0
    lsImpre = ""
    
    lsFechaIni = Mid(psFechaIni, 7, 4) & Mid(psFechaIni, 4, 2) & Mid(psFechaIni, 1, 2)
    lsFechaFin = Mid(psFechaFin, 7, 4) & Mid(psFechaFin, 4, 2) & Mid(psFechaFin, 1, 2)
        
    Set oCon = New DConecta
    oCon.AbreConexion

    'Sql = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0 ORDER BY dFecCamb"

    SQL = "SELECT TOP 1 nValFijo FROM TIPOCAMBIO WHERE datediff(d, dateadd(d, 1 , '" _
        & Format(psFechaFin, gsFormatoFecha) & "'), dFecCamb) = 0 ORDER BY dFecCamb"

    Set rs = oCon.CargaRecordSet(SQL)
        If Not (rs.EOF And rs.BOF) Then
            lnTipoCambio = rs!nValFijo
        Else
            lnTipoCambio = 0
        End If
    Set rs = Nothing
      
    lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS MENSUALES EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
    lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
    lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
    
    lnContLin = 4
    
    'Sql = " SELECT C.cPersCod, cPersNombre, cPersIDNro, Monto FROM " _
        & " (SELECT cPersCod, SUM(MontoS/TipoCambio + MontoD) AS Monto FROM " _
        & " (SELECT CV.cPersCod, MontoS = 0, nMovImporte MontoD, " _
        & " TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCOMPRAVENTA CV INNER JOIN MOV M ON CV.nMovNro = M.nMovNro " _
        & " WHERE SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 "
       
       SQL = " SELECT C.cPersCod, cPersNombre, cPersIDNro, Monto FROM " _
        & " (SELECT cPersCod, SUM(MontoS/TipoCambio + MontoD) AS Monto FROM " _
        & " (SELECT CV.cPersCod, MontoS = 0, nMovImporte MontoD, " _
        & " TipoCambio = (SELECT TOP 1 nValFijo FROM TIPOCAMBIO WHERE datediff(d, dateadd(d,1 , '" _
        & Format(psFechaFin, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCOMPRAVENTA CV INNER JOIN MOV M ON CV.nMovNro = M.nMovNro " _
        & " WHERE SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 "
   
   If psCodAge <> "" Then
        SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
   End If
   
'   sql = sql & "UNION All " _
        & " SELECT cPersCod = (SELECT TOP 1 cPersCod FROM PRODUCTOPERSONA PP WHERE PP.cCtaCod = MC.cCtaCod ORDER BY cPersCod), " _
        & " MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & " MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & " TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCAP MC INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 AND " _
        & " M.cOpeCod in (" _
        & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
   
   SQL = SQL & "UNION All " _
        & " SELECT cPersCod, " _
        & " MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & " MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & " TipoCambio = (SELECT TOP 1 nValFijo FROM TIPOCAMBIO WHERE datediff(d, dateadd(d,1, '" _
        & Format(psFechaFin, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCAP MC " _
        & " inner join productopersona pp on MC.cctacod=pp.cctacod " _
        & " INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 AND pp.nprdPersRelac=10 and " _
        & " M.cOpeCod not in (401401, 402401) and not M.cOpeCod like '9010%'"
        
    If psCodAge <> "" Then
         SQL = SQL & " AND SUBSTRING(MC.cCtaCod, 4,2) = '" & psCodAge & "'"
    End If
        
    'Sql = Sql & "UNION All " _
        & "SELECT cPersCod, " _
        & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & "TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & "FROM MOVCOL MC " _
        & " inner join productopersona pp on MC.cctacod=pp.cctacod " _
        & " INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 AND pp.nprdPersRelac=20 AND " _
        & " M.cOpeCod in (" _
        & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
        
    SQL = SQL & "UNION All " _
        & "SELECT cPersCod, " _
        & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & "TipoCambio = (SELECT TOP 1 nValFijo FROM TIPOCAMBIO WHERE datediff(d, dateadd(d, 1, '" _
        & Format(psFechaFin, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & "FROM MOVCOL MC " _
        & " inner join productopersona pp on MC.cctacod=pp.cctacod " _
        & " INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) BETWEEN '" _
        & lsFechaIni & "' AND '" & lsFechaFin & "' AND M.nMovFlag = 0 AND pp.nprdPersRelac=20 AND " _
        & " M.cOpeCod in (" _
        & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
                   
    If psCodAge <> "" Then
         SQL = SQL & " AND SUBSTRING(MC.cCtaCod, 4,2) = '" & psCodAge & "'"
    End If
        
    SQL = SQL & ") S " _
        & "GROUP BY cPersCod) C " _
        & "INNER JOIN PERSONA P On C.cPersCod = P.cPersCod INNER JOIN PERSID DI ON DI.cPersCod = C.cPersCod " _
        & "WHERE Monto >= " & lnMontoMenLD & " AND cPersIDTpo in (1,2) "
        
     'Sql = Sql & ") S " _
        & "GROUP BY cPersCod) C " _
        & "INNER JOIN PERSONA P On C.cPersCod = P.cPersCod " _
        & "WHERE Monto >= " & lnMontoMenLD & ""
        
        
    Set rs = oCon.CargaRecordSet(SQL)

    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
                
            If lnContLin > 52 Then
                lnContPag = lnContPag + 1
                lsImpre = lsImpre & Chr(12)
                lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
                lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS MENSUALES EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
                lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
                lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
                lnContLin = 4
            End If
            
            lsCliente = PstaNombre(rs!cPersNombre)
            lsDocId = rs!cPersIDNro
            
            lsImpre = lsImpre & rs!cPersCod & Space(3) & ImpreCarEsp(lsCliente) & Space(5) & lsDocId & Chr(10)
            
            lsImpre = lsImpre & String(180, "-") & Chr(10)
            lsImpre = lsImpre & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
            lsImpre = lsImpre & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
            lsImpre = lsImpre & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
            lsImpre = lsImpre & String(180, "-") & Chr(10)
            
            lnContLin = lnContLin + 4
            
            SQL = " SELECT nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, Agencia, cMovNro, TipoProd, cOpeDesc, " _
                & " cCtaCod, MontoS, MontoD, SUBSTRING(cMovNro, 22,4) Usuario FROM (" _
                & " SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
                & " TipoProd = '', cOpeDesc, cCtaCod = '', MontoS = 0, nMovImporte MontoD, SUBSTRING(cMovNro, 22,4) Usuario  " _
                & " FROM MOVCOMPRAVENTA CV INNER JOIN PERSONA P ON CV.cPersCod = P.cPersCod " _
                & " INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = CV.nMovNro " _
                & " INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod " _
                & " INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & " WHERE cPersIDTpo in (1,2) AND CV.cPersCod = '" & rs!cPersCod & "' AND SUBSTRING(cMovNro,1,8) BETWEEN '" _
                & lsFechaIni & "' AND '" & lsFechaFin & "' "
                
            If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "' "
            End If
                
            'sql = sql & "UNION All " _
                & "SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
                & "TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 232 THEN 'AHORROS' WHEN 233 " _
                & "THEN 'PLAZOFIJO' WHEN 234 THEN 'CTS' END, cOpeDesc, MC.cCtaCod, " _
                & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
                & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
                & "FROM MOVCAP MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
                & "PERSONA P ON PP.cPersCod = P.cPersCod " _
                & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
                & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & "WHERE cPersIDTpo in (1,2) AND PP.cPersCod = '" & Rs!cPersCod _
                & "' AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "' AND " _
                & " M.cOpeCod in (" _
                & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
                & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
                & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
                
             SQL = SQL & "UNION All " _
                & "SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
                & "TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 232 THEN 'AHORROS' WHEN 233 " _
                & "THEN 'PLAZOFIJO' WHEN 234 THEN 'CTS' END, cOpeDesc, MC.cCtaCod, " _
                & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
                & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
                & "FROM MOVCAP MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
                & "PERSONA P ON PP.cPersCod = P.cPersCod " _
                & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
                & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & "WHERE cPersIDTpo in (1,2) AND PP.cPersCod = '" & rs!cPersCod _
                & "' AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "' AND pp.nprdPersRelac=10 AND " _
                & " MC.cOpeCod not in (401401, 402401) and not MC.cOpeCod like '9010%' and m.nmovflag=0"
                                
             If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(MC.cCtaCod, 4,2) = '" & psCodAge & "'"
             End If
                                
             SQL = SQL & "UNION All " _
                & "SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, " _
                & " cMovNro, TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 305 THEN 'PIGNORATICIO' ELSE 'CREDITOS' END, " _
                & " cOpeDesc, MC.cCtaCod, MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
                & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
                & "FROM MOVCOL MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
                & "PERSONA P ON PP.cPersCod = P.cPersCod " _
                & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
                & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & "WHERE cPersIDTpo in (1,2) AND PP.cPersCod = '" & rs!cPersCod _
                & "' AND SUBSTRING(cMovNro,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "' AND nPrdPersRelac = 20 AND " _
                & " M.cOpeCod in (" _
                & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
                & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
                & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
                
                                           
             If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(MC.cCtaCod, 4,2) = '" & psCodAge & "'"
             End If
                                
           SQL = SQL & ") S ORDER BY cMovNro"
           
            Set rsDet = oCon.CargaRecordSet(SQL)
            lnItem = 0
            lnTCItem = 0
            lnTCMontoS = 0
            lnTCMontoED = 0
            lnTCMontoD = 0
            
            If Not rsDet.EOF And Not rs.BOF Then
                Do While Not rsDet.EOF
                
                    If lnContLin >= 60 Then
                        lnContPag = lnContPag + 1
                        lsImpre = lsImpre & Chr(12)
                        
                        lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
                        lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS MENSUALES EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
                        lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
                        lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
                        
                        lsImpre = lsImpre & String(180, "-") & Chr(10)
                        lsImpre = lsImpre & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
                        lsImpre = lsImpre & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
                        lsImpre = lsImpre & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
                        lsImpre = lsImpre & String(180, "-") & Chr(10)
                
                        lnContLin = 7
                    End If
                    
                    lnItem = lnItem + 1
                    lsAgencia = rsDet!cAgeDescripcion
                    lsFechaTrans = Mid(rsDet!cMovNro, 7, 2) & "/" & Mid(rsDet!cMovNro, 5, 2) & "/" & Mid(rsDet!cMovNro, 1, 4) & " " & Mid(rsDet!cMovNro, 9, 2) & ":" & Mid(rsDet!cMovNro, 11, 2) & ":" & Mid(rsDet!cMovNro, 13, 2)
                    lsProducto = rsDet!TipoProd
                    lsOperacion = rsDet!cOpeDesc
                    lsCtaCod = rsDet!Cctacod
                    lnMontoED = Format(rsDet!MontoS / lnTipoCambio, "#0.00")
                    
                    lsImpre = lsImpre & ImpreFormat(lnItem, 3, 0) & Space(3) & lsAgencia & Space(5) & lsFechaTrans
                    lsImpre = lsImpre & Space(5) & lsProducto & Space(5) & lsOperacion & Space(5) & lsCtaCod & Space(5)
                    lsImpre = lsImpre & ImpreFormat(rsDet!MontoS, 11, 2, True) & Space(3) & ImpreFormat(lnMontoED, 11, 2, True)
                    lsImpre = lsImpre & Space(3) & ImpreFormat(rsDet!MontoD, 11, 2) & Space(3)
                    lsImpre = lsImpre & ImpreFormat(lnMontoED + rsDet!MontoD, 11, 2) & Space(3) & rsDet!Usuario & Chr(10)
                    
                    lnContLin = lnContLin + 1
                    
                    lnTCItem = lnTCItem + 1
                    lnTCMontoS = lnTCMontoS + rsDet!MontoS
                    lnTCMontoED = lnTCMontoED + lnMontoED
                    lnTCMontoD = lnTCMontoD + rsDet!MontoD
                
                    rsDet.MoveNext
                Loop
            End If
            
            lsImpre = lsImpre & String(181, "-") & Chr(10)
            lsImpre = lsImpre & Space(5) & "TOTAL CLIENTE  :  " & ImpreFormat(lnTCItem, 3, 0) & Space(80) & ImpreFormat(lnTCMontoS, 13, 2, True) & Space(1)
            lsImpre = lsImpre & ImpreFormat(lnTCMontoED, 13, 2, True) & Space(1) & ImpreFormat(lnTCMontoD, 13, 2, True) & Space(1)
            lsImpre = lsImpre & ImpreFormat(lnTCMontoED + lnTCMontoD, 13, 2, True) & Chr(10)
            lsImpre = lsImpre & String(181, "-") & Chr(10) & Chr(10)
            
            lnContLin = lnContLin + 3
            
            lnTItem = lnTItem + lnTCItem
            lnTMontoS = lnTMontoS + lnTCMontoS
            lnTMontoED = lnTMontoED + lnTCMontoED
            lnTMontoD = lnTMontoD + lnTCMontoD
            
            Set rsDet = Nothing
            rs.MoveNext
        Loop
    End If

    lsImpre = lsImpre & String(181, "-") & Chr(10)
    lsImpre = lsImpre & Space(5) & "TOTAL          :  " & ImpreFormat(lnTItem, 3, 0) & Space(80) & ImpreFormat(lnTMontoS, 13, 2, True) & Space(1)
    lsImpre = lsImpre & ImpreFormat(lnTMontoED, 13, 2, True) & Space(1) & ImpreFormat(lnTMontoD, 13, 2, True) & Space(1)
    lsImpre = lsImpre & ImpreFormat(lnTMontoED + lnTMontoD, 13, 2, True) & Chr(10)
    lsImpre = lsImpre & String(181, "-") & Chr(10)
    
    Set rs = Nothing
     
    ImpLavDineroM = lsImpre
    
End Function

Public Function ImpLavDineroAD(ByVal psFechaIni As String, ByVal psNomCmac As String, _
                              ByVal psNomAge As String, ByVal psCodAge As String, pdFecSis As Date) As String

Dim SQL As String
Dim oCon As DConecta
Dim rs As Recordset
Dim rsDet As Recordset
Dim lsImpre As String
Dim oGen As DGeneral
Dim lnMontoMenLD As Currency
Dim lsCliente As String * 40
Dim lnMontoED As Currency
Dim lnTCMontoS As Currency
Dim lnTCMontoED As Currency
Dim lnTCMontoD As Currency
Dim lnTMontoS As Currency
Dim lnTMontoED As Currency
Dim lnTMontoD As Currency
Dim lnItem As Long, lnTCItem As Long, lnTItem As Long
Dim lsDocId As String * 8
Dim lsAgencia As String * 15, lsFechaTrans As String * 19
Dim lsCtaCod As String * 18, lsFechas As String * 10
Dim lsProducto As String * 10, lsOperacion As String * 15
Dim lnTipoCambio As Currency
Dim lnContPag As Integer
Dim lnContLin As Integer
Dim lsFechaIni As String, lsFechaFin As String

    Set oGen = New DGeneral
        lnMontoMenLD = oGen.GetParametro(2000, 2032)
    Set oGen = Nothing

    lnContPag = 1:  lnItem = 0
    lnTCItem = 0: lnTItem = 0
    lsImpre = ""
    
    lsFechaIni = Mid(psFechaIni, 7, 4) & Mid(psFechaIni, 4, 2) & Mid(psFechaIni, 1, 2)
        
    Set oCon = New DConecta
    oCon.AbreConexion

    SQL = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0 ORDER BY dFecCamb"

    Set rs = oCon.CargaRecordSet(SQL)
        If Not (rs.EOF And rs.BOF) Then
            lnTipoCambio = rs!nValPond
        Else
            lnTipoCambio = 0
        End If
    Set rs = Nothing
        
    lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS DIARIAS EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
    lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
    lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
    
    lnContLin = 4
    lnContPag = 1
    
    SQL = " SELECT C.cPersCod, cPersNombre, cPersIDNro, Monto FROM " _
        & " (SELECT cPersCod, SUM(MontoS/TipoCambio + MontoD) AS Monto FROM " _
        & " (SELECT CV.cPersCod, MontoS = 0, nMovImporte MontoD, " _
        & " TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCOMPRAVENTA CV INNER JOIN MOV M ON CV.nMovNro = M.nMovNro " _
        & " WHERE SUBSTRING(cMovNro,1,8) = '" _
        & lsFechaIni & "' AND M.nMovFlag = 0 "
        
   If psCodAge <> "" Then
        SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
   End If
   
   SQL = SQL & "UNION All " _
        & " SELECT cPersCod = (SELECT TOP 1 cPersCod FROM PRODUCTOPERSONA PP WHERE PP.cCtaCod = MC.cCtaCod ORDER BY cPersCod), " _
        & " MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & " MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & " TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & " FROM MOVCAP MC INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) = '" _
        & lsFechaIni & "' AND M.nMovFlag = 0 AND " _
        & " M.cOpeCod in (" _
        & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
        
    If psCodAge <> "" Then
         SQL = SQL & " AND SUBSTRING(M.cMovNro, 18,2) = '" & psCodAge & "'"
    End If
        
   SQL = SQL & "UNION All " _
        & "SELECT cPersCod = (SELECT TOP 1 cPersCod FROM PRODUCTOPERSONA PP WHERE PP.cCtaCod = MC.cCtaCod " _
        & " And nPrdPersRelac = 20 ORDER BY cPersCod), " _
        & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
        & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, " _
        & "TipoCambio = (SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(psFechaIni, 1, 2) * -1 & ", '" _
        & Format(psFechaIni, gsFormatoFecha) & "'), dFecCamb) = 0) " _
        & "FROM MOVCOL MC INNER JOIN MOV M ON MC.nMovNro = M.nMovNro AND SUBSTRING(cMovNro,1,8) = '" _
        & lsFechaIni & "' AND M.nMovFlag = 0 AND " _
        & " M.cOpeCod in (" _
        & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
        
    If psCodAge <> "" Then
         SQL = SQL & " AND SUBSTRING(M.cMovNro, 18,2) = '" & psCodAge & "'"
    End If
        
    SQL = SQL & ") S " _
        & "GROUP BY cPersCod) C " _
        & "INNER JOIN PERSONA P On C.cPersCod = P.cPersCod INNER JOIN PERSID DI ON DI.cPersCod = C.cPersCod " _
        & "WHERE Monto >= " & lnMontoMenLD & " AND cPersIDTpo in (1,2) "
        
    Set rs = oCon.CargaRecordSet(SQL)

    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
                
            If lnContLin > 52 Then
                lnContPag = lnContPag + 1
                lsImpre = lsImpre & Chr(12)
                lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
                lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS DIARIAS EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
                lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
                lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
                lnContLin = 4
            End If
            
            lsCliente = PstaNombre(rs!cPersNombre)
            lsDocId = rs!cPersIDNro
            
            lsImpre = lsImpre & rs!cPersCod & Space(3) & ImpreCarEsp(lsCliente) & Space(5) & lsDocId & Chr(10)
            
            lsImpre = lsImpre & String(180, "-") & Chr(10)
            lsImpre = lsImpre & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
            lsImpre = lsImpre & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
            lsImpre = lsImpre & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
            lsImpre = lsImpre & String(180, "-") & Chr(10)
            
            lnContLin = lnContLin + 4
            
            SQL = " SELECT nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, Agencia, cMovNro, TipoProd, cOpeDesc, " _
                & " cCtaCod, MontoS, MontoD, SUBSTRING(cMovNro, 22,4) Usuario FROM (" _
                & " SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
                & " TipoProd = '', cOpeDesc, cCtaCod = '', MontoS = 0, nMovImporte MontoD, SUBSTRING(cMovNro, 22,4) Usuario  " _
                & " FROM MOVCOMPRAVENTA CV INNER JOIN PERSONA P ON CV.cPersCod = P.cPersCod " _
                & " INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = CV.nMovNro " _
                & " INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod " _
                & " INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & " WHERE cPersIDTpo in (1,2) AND CV.cPersCod = '" & rs!cPersCod & "' AND SUBSTRING(cMovNro,1,8) = '" _
                & lsFechaIni & "' "
                
            If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "' "
            End If
                
            SQL = SQL & "UNION All " _
                & "SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, cMovNro, " _
                & "TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 232 THEN 'AHORROS' WHEN 233 " _
                & "THEN 'PLAZOFIJO' WHEN 234 THEN 'CTS' END, cOpeDesc, MC.cCtaCod, " _
                & "MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
                & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
                & "FROM MOVCAP MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
                & "PERSONA P ON PP.cPersCod = P.cPersCod " _
                & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
                & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & "WHERE cPersIDTpo in (1,2) AND PP.cPersCod = '" & rs!cPersCod _
                & "' AND SUBSTRING(cMovNro,1,8) = '" & lsFechaIni & "' AND " _
                & " M.cOpeCod in (" _
                & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
                & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
                & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
                
             If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
             End If
                                
            SQL = SQL & "UNION All " _
                & "SELECT M.nMovNro, cPersNombre, cPersIDNro, cAgeDescripcion, SUBSTRING(cMovNro, 18,2) Agencia, " _
                & " cMovNro, TipoProd = CASE SUBSTRING(MC.cCtaCod,6,3) WHEN 305 THEN 'PIGNORATICIO' ELSE 'CREDITOS' END, " _
                & " cOpeDesc, MC.cCtaCod, MontoS = CASE substring(MC.cCtaCod, 9,1) WHEN 1 THEN nMonto ELSE 0 END, " _
                & "MontoD = CASE substring(MC.cCtaCod, 9,1) WHEN 2 THEN nMonto ELSE 0 END, SUBSTRING(cMovNro, 22,4) Usuario " _
                & "FROM MOVCOL MC INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = MC.cCtaCod INNER JOIN " _
                & "PERSONA P ON PP.cPersCod = P.cPersCod " _
                & "INNER JOIN PersID DI ON P.cPersCod = DI.cPersCod INNER JOIN MOV M ON M.nMovNro = MC.nMovNro " _
                & "INNER JOIN OPETPO O ON M.cOpeCod = O.cOpeCod INNER JOIN AGENCIAS A ON A.cAgeCod = SUBSTRING(cMovNro, 18,2) " _
                & "WHERE cPersIDTpo in (1,2) AND PP.cPersCod = '" & rs!cPersCod _
                & "' AND SUBSTRING(cMovNro,1,8) = '" & lsFechaIni & "' AND nPrdPersRelac = 20 AND " _
                & " M.cOpeCod in (" _
                & " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
                & " WHERE G.nEfectivo = 1 and not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
                & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47)) "
                
             If psCodAge <> "" Then
                SQL = SQL & " AND SUBSTRING(cMovNro, 18,2) = '" & psCodAge & "'"
             End If
                                
           SQL = SQL & ") S ORDER BY cMovNro"
           
            Set rsDet = oCon.CargaRecordSet(SQL)
            lnItem = 0
            lnTCItem = 0
            lnTCMontoS = 0
            lnTCMontoED = 0
            lnTCMontoD = 0
            
            If Not rsDet.EOF And Not rs.BOF Then
                Do While Not rsDet.EOF
                
                    If lnContLin >= 60 Then
                        lnContPag = lnContPag + 1
                        lsImpre = lsImpre & Chr(12)
                        
                        lsImpre = lsImpre & psNomCmac & Space(125) & "Fecha : " & Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss") & Chr(10)
                        lsImpre = lsImpre & psNomAge & Space(18) & "OPERACIONES ACUMULADAS DIARIAS EN EFECTIVO MAYORES O EQUIVALENTES A : " & lnMontoMenLD & "  DOLARES"
                        lsImpre = lsImpre & Space(46) & "Pag. " & lnContPag & Chr(10)
                        lsImpre = lsImpre & Space(80) & "TIPO DE CAMBIO :   " & lnTipoCambio & Chr(10) & Chr(10)
                        
                        lsImpre = lsImpre & String(180, "-") & Chr(10)
                        lsImpre = lsImpre & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
                        lsImpre = lsImpre & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
                        lsImpre = lsImpre & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
                        lsImpre = lsImpre & String(180, "-") & Chr(10)
                
                        lnContLin = 7
                    End If
                    
                    lnItem = lnItem + 1
                    lsAgencia = rsDet!cAgeDescripcion
                    lsFechaTrans = Mid(rsDet!cMovNro, 7, 2) & "/" & Mid(rsDet!cMovNro, 5, 2) & "/" & Mid(rsDet!cMovNro, 1, 4) & " " & Mid(rsDet!cMovNro, 9, 2) & ":" & Mid(rsDet!cMovNro, 11, 2) & ":" & Mid(rsDet!cMovNro, 13, 2)
                    lsProducto = rsDet!TipoProd
                    lsOperacion = rsDet!cOpeDesc
                    lsCtaCod = rsDet!Cctacod
                    lnMontoED = Format(rsDet!MontoS / lnTipoCambio, "#0.00")
                    
                    lsImpre = lsImpre & ImpreFormat(lnItem, 3, 0) & Space(3) & lsAgencia & Space(5) & lsFechaTrans
                    lsImpre = lsImpre & Space(5) & lsProducto & Space(5) & lsOperacion & Space(5) & lsCtaCod & Space(5)
                    lsImpre = lsImpre & ImpreFormat(rsDet!MontoS, 11, 2, True) & Space(3) & ImpreFormat(lnMontoED, 11, 2, True)
                    lsImpre = lsImpre & Space(3) & ImpreFormat(rsDet!MontoD, 11, 2) & Space(3)
                    lsImpre = lsImpre & ImpreFormat(lnMontoED + rsDet!MontoD, 11, 2) & Space(3) & rsDet!Usuario & Chr(10)
                    
                    lnContLin = lnContLin + 1
                    
                    lnTCItem = lnTCItem + 1
                    lnTCMontoS = lnTCMontoS + rsDet!MontoS
                    lnTCMontoED = lnTCMontoED + lnMontoED
                    lnTCMontoD = lnTCMontoD + rsDet!MontoD
                
                    rsDet.MoveNext
                Loop
            End If
            
            lsImpre = lsImpre & String(181, "-") & Chr(10)
            lsImpre = lsImpre & Space(5) & "TOTAL CLIENTE  :  " & ImpreFormat(lnTCItem, 3, 0) & Space(80) & ImpreFormat(lnTCMontoS, 13, 2, True) & Space(1)
            lsImpre = lsImpre & ImpreFormat(lnTCMontoED, 13, 2, True) & Space(1) & ImpreFormat(lnTCMontoD, 13, 2, True) & Space(1)
            lsImpre = lsImpre & ImpreFormat(lnTCMontoED + lnTCMontoD, 13, 2, True) & Chr(10)
            lsImpre = lsImpre & String(181, "-") & Chr(10) & Chr(10)
            
            lnContLin = lnContLin + 3
            
            lnTItem = lnTItem + lnTCItem
            lnTMontoS = lnTMontoS + lnTCMontoS
            lnTMontoED = lnTMontoED + lnTCMontoED
            lnTMontoD = lnTMontoD + lnTCMontoD
            
            Set rsDet = Nothing
            rs.MoveNext
        Loop
    End If

    lsImpre = lsImpre & String(181, "-") & Chr(10)
    lsImpre = lsImpre & Space(5) & "TOTAL          :  " & ImpreFormat(lnTItem, 3, 0) & Space(80) & ImpreFormat(lnTMontoS, 13, 2, True) & Space(1)
    lsImpre = lsImpre & ImpreFormat(lnTMontoED, 13, 2, True) & Space(1) & ImpreFormat(lnTMontoD, 13, 2, True) & Space(1)
    lsImpre = lsImpre & ImpreFormat(lnTMontoED + lnTMontoD, 13, 2, True) & Chr(10)
    lsImpre = lsImpre & String(181, "-") & Chr(10)
    
    Set rs = Nothing
     
    ImpLavDineroAD = lsImpre
    
End Function


Public Function SobranteFaltante(pdFecIni As String, pdFecFin As String, psAge As String) As String
    Dim i As Integer
    Dim rsSobFal As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsRTFExt As String
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And SubString(M.cMovNro,18,2) = '" & psAge & "'"
    End If
    
    oCon.AbreConexion
    
    If pdFecIni = pdFecFin Then
        pdFecFin = Format$(DateAdd("d", 1, CDate(pdFecFin)), gsFormatoFechaView)
    End If
    lsRTFExt = ""
    'Sobrantes de caja
    For i = 1 To 2
        VSQL = "Select dbo.GetFechaMov(M.cMovNro,103) dFecTran, MOV.nMovImporte nMonTran, IsNull(PE.cPersNombre,'') cNomUsu, Right(M.cMovNro,4) cCodUsu " _
             & " From Mov M" _
             & " Inner Join MovOpeVarias MOV On M.nMovNro = MOV.nMovNro" _
             & " Left JOIN RRHH RH ON Right(M.cMovNro,4) = RH.cUser" _
             & " Left JOIN Persona PE ON RH.cPersCod = PE.cPersCod" _
             & " where (Left(M.cMovNro,8) between '" & Format$(pdFecIni, gsFormatoMovFecha) & "' and '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "') and M.cOpeCod in ('" & gOpeHabCajRegSobrante & "')" _
             & " And MOV.nMoneda = " & i & " And M.nMovFlag = 0 " & sqlAdd & "   " _
             & " Order by M.cMovNro, cCodUsu"
        If rsSobFal.State = adStateOpen Then rsSobFal.Close
        Set rsSobFal = oCon.CargaRecordSet(VSQL)
        If RSVacio(rsSobFal) Then
            lsRTFExt = lsRTFExt & ""
        Else
            If Len(Trim(lsRTFExt)) <> 0 Then
                lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoPagina
            End If
            lsRTFExt = lsRTFExt & PrtSobFal(rsSobFal, i, True)
        End If
    Next i
    
    'Faltantes de Caja
    For i = 1 To 2
        VSQL = "Select dbo.GetFechaMov(M.cMovNro,103) dFecTran, MOV.nMovImporte nMonTran, IsNull(PE.cPersNombre,'') cNomUsu, Right(M.cMovNro,4) cCodUsu " _
             & " From Mov M" _
             & " Inner Join MovOpeVarias MOV On M.nMovNro = MOV.nMovNro" _
             & " Left JOIN RRHH RH ON Right(M.cMovNro,4) = RH.cUser" _
             & " Left JOIN Persona PE ON RH.cPersCod = PE.cPersCod" _
             & " where (Left(M.cMovNro,8) between '" & Format$(pdFecIni, gsFormatoMovFecha) & "' and '" & Format$(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "') and M.cOpeCod in ('" & gOpeHabCajRegFaltante & "')" _
             & " And MOV.nMoneda = " & i & " And M.nMovFlag = 0  " & sqlAdd & "    " _
             & " Order by dFectran, cCodUsu"
        If rsSobFal.State = adStateOpen Then rsSobFal.Close
        Set rsSobFal = oCon.CargaRecordSet(VSQL)
        If RSVacio(rsSobFal) Then
            lsRTFExt = lsRTFExt & ""
        Else
            If Len(Trim(lsRTFExt)) <> 0 Then
                lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoPagina
            End If
            lsRTFExt = lsRTFExt & PrtSobFal(rsSobFal, i, False)
        End If
    Next i
    
    Set rsSobFal = Nothing
    SobranteFaltante = lsRTFExt
End Function


Private Function PrtSobFal(prSobFal As ADODB.Recordset, pnMoneda As Integer, Optional pnTipo As Boolean = True) As String
Dim lnAcum As Currency
Dim lnCarLin As Integer
Dim lsTitRp1 As String
Dim lsTitRp2 As String
Dim lsMoneda As String
Dim lsNumPag As String
Dim lsFchtra As String
Dim lnLinPag As Integer
Dim lnCntPag As Integer
Dim lsNomOpe As String
Dim lsUsu As String * 35
Dim lsRTFExt As String
lsNumPag = ""
lnLinPag = 0
lnCarLin = 85
lsTitRp1 = IIf(pnTipo, "L I S T A D O   D E   S O B R A N T E S ", "L I S T A D O   D E   F A L T A N T E S")
lsTitRp2 = "" '"USO DE CONTABILIDAD"
lnCntPag = 1
lnAcum = 0
lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
With prSobFal
    lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
    lsRTFExt = lsRTFExt & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "  FECHA OPERACION         USU.  NOMBRE                                      MONTO      " & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lnLinPag = 8
    Do While Not .EOF
        lnAcum = lnAcum + !nMonTran
        lsUsu = PstaNombre(Trim(!cNomUsu), True)
        lsRTFExt = lsRTFExt & Format$(!dFecTran, "dd/mm/yyyy hh:mm:ss AMPM") & "    " & !cCodUsu & "  " & lsUsu & "  " & JDNum(Trim(Str(IIf(IsNull(!nMonTran), 0, !nMonTran))), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lnLinPag = lnLinPag + 1
        If lnLinPag > 58 Then
            lsRTFExt = lsRTFExt & gPrnSaltoPagina
            lnCntPag = lnCntPag + 1
            lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
            lsRTFExt = lsRTFExt & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
            lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
            lsRTFExt = lsRTFExt & "  FECHA OPERACION         USU.  NOMBRE                                      MONTO      " & oImpresora.gPrnSaltoLinea
            lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
            lnLinPag = 8
        End If
        .MoveNext
    Loop
    lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(40, " ") & "******** ACUMULADO " & FillText(lsMoneda, 8, " ") & "  " & JDNum(Trim(Str(lnAcum)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
End With

PrtSobFal = lsRTFExt
End Function


Function RepoCheques(ByVal FechaI As Date, ByVal FechaF As Date) As String
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Age(50) As String
Dim ValAge(50) As String
Dim i As Long
Dim J As Byte
Dim lsCadena As String
Dim vLineas As Integer
Dim Pag As Integer
Dim SubTotal As Currency
Dim Total As Currency
Dim Usuario As String
Dim UsuarioR As String
Dim lbCabecera As Boolean
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

Dim lsCtaBco As String * 11
Dim lsNroCheque As String * 13
Dim lsPlaza As String * 11
Dim lsBco As String * 30

Dim sqlAdd As String

If gsCodAge <> "" Then
    
    sqlAdd = " And Substring(M.cMovNro,18,2) = '" & gsCodAge & "' "
    
    SQL = "SELECT cAgeDescripcion cNomTab, cAgeCod cValor FROM Agencias Where cAgeCod = '" & gsCodAge & "'"
    Set rs = oCon.CargaRecordSet(SQL)
    If RSVacio(rs) Then
    Else
       Do While Not rs.EOF
       Age(1) = rs!cNomTab
       ValAge(1) = rs!cValor
       rs.MoveNext
       Loop
    End If
    rs.Close

    SQL = "SELECT cAgeDescripcion cNomTab, cAgeCod cValor FROM Agencias Where cAgeCod <> '" & gsCodAge & "'"
    Set rs = oCon.CargaRecordSet(SQL)
    If RSVacio(rs) Then
    Else
    i = 2
       Do While Not rs.EOF
       Age(i) = rs!cNomTab
       ValAge(i) = rs!cValor
       rs.MoveNext
       i = i + 1
       Loop
    End If
    rs.Close
Else
    SQL = "SELECT cAgeDescripcion cNomTab, cAgeCod cValor FROM Agencias"
    Set rs = oCon.CargaRecordSet(SQL)
    If RSVacio(rs) Then
    Else
       i = 1
       Do While Not rs.EOF
       Age(i) = rs!cNomTab
       ValAge(i) = rs!cValor
       rs.MoveNext
       i = i + 1
       Loop
    End If
    rs.Close
End If
SubTotal = 0
Total = 0
vLineas = 8
Pag = 1
'*************  REPORTE ********************
Dim lsMoneda As String
For J = 1 To 2
    lbCabecera = False
    Total = 0
    SQL = " SELECT  IsNull((Select cPersNombre From Persona PE where PE.cPersCod = DR.cPersCod),'') Bco, DRC.cCtaCod cCodCta, DRC.cNroDoc cNumChq, dbo.FechaMov(M.cMovNro) dRegChq," _
        & " DR.dValorizacion dValorChq, DR.nMonto nMontoChq, Right(M.cMovNro,4) cCodUsu," _
        & " M.cOpeCod cCodOpe, CO.cConsDescripcion cNomTab, DR.cIFCta cCtaBco, SubString(M.cMovNro,18,2) cCodAge, IsNull((Select cAgeDescripcion From Agencias AGE Where AGE.cAgeCod = SubString(M.cMovNro,18,2)),'') Agencia " _
        & " FROM DOcRecCapta DRC" _
        & " Inner Join DocRec DR On DRC.nTpoDoc = DR.nTpoDoc And DRC.cNroDoc = DR.cNroDoc" _
        & "                         And DRC.cPersCod = DR.cPersCod And DRC.cIFTpo = DR.cIFTpo" _
        & " Inner Join Mov M On M.nMovNro = DRC.nMovNro And M.cOpeCod in ('" & gChqOpeRegistro & "','" & gAhoApeChq & "','" & gAhoApeLoteChq & "','" & gAhoDepChq & "','" & gPFApeChq & "','" & gPFApeLoteChq & "','" & gCTSApeChq & "','" & gCTSApeLoteChq & "','" & gCTSDepChq & "')" _
        & "                         And Left(M.cMovNro,8) Between '" & Format(FechaI, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, FechaF), gsFormatoMovFecha) & "'" _
        & "                         And Substring(DRC.cCtaCod,9,1)= '" & J & "'  and DR.nEstado not in (" & ChequeEstado.gChqEstAnulado & "," & ChequeEstado.gChqEstRechazado & "," & ChequeEstado.gsChqEstExtornado & ")" _
        & "                         And M.nMovflag = 0 And DRC.nTpoDoc = '" & TpoDoc.TpoDocCheque & "' " & sqlAdd & "" _
        & " Inner Join Constante CO On CO.nConsCod = " & gChequePlaza & " And CO.nConsvalor = DR.bPlaza"
        Set rs = oCon.CargaRecordSet(SQL)
        If RSVacio(rs) Then
        Else
            If Not lbCabecera Then
                lsMoneda = IIf(J = 1, "SOLES", "DOLARES")
                lsCadena = lsCadena + UCase(gsNomAge) & " - " & lsMoneda & Space(78) + " " & ArmaFecha(gdFecSis) & " - " & Format(Now, "HH:MM:SS AMPM") & oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena + Space(40) + "LISTADO DE CHEQUES RECIBIDOS" & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena + String(154, "=") + oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena + "CUENTA              " + Space(2) & "NUMCHQ  " + Space(2) & "   BANCO                         "
                lsCadena = lsCadena + Space(2) + "        MONTO" + Space(2) + "FECHA REG "
                lsCadena = lsCadena + Space(2) & "FECHA VAL." + Space(2) + "USUA"
                lsCadena = lsCadena + Space(2) & "CUENTA      "
                lsCadena = lsCadena + Space(2) & "PLAZA       "
                lsCadena = lsCadena + Space(2) & "AGENCIA " & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena + String(154, "=")
                lbCabecera = True
            End If
            lsCadena = lsCadena + oImpresora.gPrnSaltoLinea + Trim(Age(i)) + oImpresora.gPrnSaltoLinea
            SubTotal = 0
            Do While Not rs.EOF
                If IsNull(rs!cCodUsu) Then
                    Usuario = "    "
                Else
                    Usuario = rs!cCodUsu
                End If
                UsuarioR = ""
                lsBco = rs!Bco
                lsNroCheque = JDNum(rs!cNumChq, 8, False, 8, 0)
                lsCadena = lsCadena & rs!cCodCta & Space(2) & lsNroCheque & Space(1)
                lsCadena = lsCadena & lsBco & Space(2) & JDNum(Str(rs!nMontoChq), 13, True, 11, 2) & Space(2)
                lsCadena = lsCadena & Format(rs!dRegChq, gsFormatoFechaView) & Space(2) & Format(rs!dValorChq, gsFormatoFechaView) + Space(2)
                lsPlaza = Trim(rs!cNomTab)
                lsCtaBco = Trim(rs!cCtaBco)
                lsCadena = lsCadena & Usuario & Space(2) & lsCtaBco & Space(2) & lsPlaza & Space(2) & rs!AGENCIA & oImpresora.gPrnSaltoLinea
                vLineas = vLineas + 1
                If vLineas > 55 Then
                    vLineas = 8
                    lsCadena = lsCadena & UCase(gsNomAge) & " - " & lsMoneda & Space(78) + " " + ArmaFecha(gdFecSis) & " - " & Format(Now, "HH:MM:SS AMPM") & oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena & Space(40) + "LISTADO DE CHEQUES RECIBIDOS" + oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena & String(154, "=") & oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena + "CUENTA              " + Space(2) & "NUMCHQ  " + Space(2) & "   BANCO                         "
                    lsCadena = lsCadena + Space(2) & "        MONTO" + Space(2) & "FECHA REG "
                    lsCadena = lsCadena + Space(2) & "FECHA VAL." + Space(2) + "USUA"
                    lsCadena = lsCadena + Space(2) & "CUENTA      "
                    lsCadena = lsCadena + Space(2) & "PLAZA       "
                    lsCadena = lsCadena + Space(2) & "AGENCIA " & oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena + String(154, "=")
                End If
                SubTotal = SubTotal + rs!nMontoChq
                rs.MoveNext
            Loop
            Total = Total + SubTotal
            lsCadena = lsCadena + Space(20) + "Sub Total : " + Space(34) + JDNum(Str(SubTotal), 13, True, 11, 2) + oImpresora.gPrnSaltoLinea
        End If
        rs.Close
    If Total > 0 Then
        lsCadena = lsCadena + oImpresora.gPrnSaltoLinea + Space(30) + "Total     : " + Space(24) + JDNum(Str(Total), 13, True, 11, 2) + oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena + oImpresora.gPrnSaltoPagina
    End If
Next J
RepoCheques = lsCadena
End Function

Public Function RepoChequesCompleto(ByVal psCodAge As String, ByVal psFecini As String, ByVal psFecFin As String, ByVal lsEstadosCheques As String, ByVal lsOptionsCheques As String, ByVal nMoneda As Integer, gsNomAge As String, gdFecSis As String) As String

Dim sCadImp As String
Dim lsSQL As String
Dim cNroDoc As String * 15
Dim cPersona As String * 25

Dim oCon As DConecta
Dim R As New ADODB.Recordset
Dim cEstadoTemp As String
Dim cAgenciaTemp As String
Dim nCantidad As Long
Dim nTotal As Double

Dim nCantidadA As Long
Dim nTotalA As Double
Set oCon = New DConecta

Dim lsCadena As String

lsCadena = "Select "
lsCadena = lsCadena & " (Select cConsDescripcion FROM Constante C Where C.nConsValor=DR.nEstado And C.nConsCod=1007) as cEstadoCheque, "
lsCadena = lsCadena & " dbo.FechaHoraMov(M.cMovNro) FechaHora, Right(M.cMovNro,4) Usu, DR.cNroDoc, "
lsCadena = lsCadena & " (Select cPersNombre FROM Persona PE Where PE.cPersCod = DR.cPersCod) Persona, DR.bPlaza, "
'agregado para sacar tipode producto en el reporte WARO 05/06/2004
lsCadena = lsCadena & "  (Select cConsDescripcion FROM Constante C Where C.nConsValor=DR.nproducto And C.nConsCod=1001) as Prod,"

lsCadena = lsCadena & " DR.cIFCta , DR.nMonto, DR.dValorizaRef, DR.dValorizacion, Substring(M.cMovNro,18,2) as cAgencia " ', DRE.nEstado, nMoneda "
lsCadena = lsCadena & " From MovDoc MD  Inner Join Mov M On M.nMovNro = MD.nMovNro "
lsCadena = lsCadena & " Inner Join DocRec DR On MD.nDocTpo = DR.nTpoDoc "
lsCadena = lsCadena & " And DR.cNroDoc = MD.cDocNro Inner Join DocRecEst DRE On DR.nTpoDoc = DRE.nTpoDoc "
lsCadena = lsCadena & " And DR.cNroDoc = DRE.cNroDoc And DR.cPersCod = DRE.cPersCod "
lsCadena = lsCadena & " And DR.cIFTpo = DRE.cIFTpo And DRE.cMovNro = M.cMovNro "

'Agregado para que saque solo cheques de captaciones
'lsCadena = lsCadena & " Inner Join DOcRecCapta DRC ON DRC.nTpoDoc = DR.nTpoDoc And DRC.cNroDoc = DR.cNroDoc And DRC.cPersCod = DR.cPersCod And DRC.cIFTpo = DR.cIFTpo "
'Fin Agregado / Actualizado por WARO ya estaba filtrado por  captaciones

lsCadena = lsCadena & " Where MD.nDocTpo = 47 And m.nMovFlag = 0 and m.copecod='900031' "

If Trim(lsOptionsCheques) = "1" Then '--Fecha de Recepcion
    lsCadena = lsCadena & " And Left(M.cMovNro,8) >='" & Format(psFecini, "YYYYMMdd") & "' And Left(M.cMovNro,8) <='" & Format(psFecFin, "YYYYMMdd") & "' "
ElseIf Trim(lsOptionsCheques) = "2" Then '--Fecha de Valorizacion
    lsCadena = lsCadena & " and convert(varchar(8), DR.dValorizacion, 112) >='" & Format(psFecini, "YYYYMMdd") & "' And convert(varchar(8), DR.dValorizacion, 112) <='" & Format(psFecFin, "YYYYMMdd") & "' "
End If
'--Estado
lsCadena = lsCadena & " and DR.nEstado in (" & lsEstadosCheques & ")"

'--Moneda
lsCadena = lsCadena & " And nMoneda=" & nMoneda
'--Agencia
If Len(Trim(psCodAge)) > 0 Then
    lsCadena = lsCadena & " And Substring(M.cMovNro,18,2) = '" & psCodAge & "' "
End If

lsCadena = lsCadena & " Order By nMoneda, DR.nEstado, Substring(M.cMovNro,18,2), M.cMovNro, Usu, DR.cNroDoc, Persona "

oCon.AbreConexion
Set R = oCon.CargaRecordSet(lsCadena)

If R.BOF Then
Else
    
    sCadImp = sCadImp & UCase(gsNomAge) & " - " & IIf(nMoneda = 1, "SOLES", "DOLARES") & Space(60) + " " + ArmaFecha(CDate(gdFecSis)) & " - " & Format(Now, "HH:MM:SS AMPM") & Chr(10) & Chr(10)
    sCadImp = sCadImp & Space(40) + "LISTADO DE CHEQUES" + Chr(10)
    sCadImp = sCadImp & String(124, "=") & Chr(10)
    sCadImp = sCadImp & Chr(10)
    
    sCadImp = sCadImp & "CHEQUES CON ESTADO: " & R!cestadocheque & Chr(10) & Chr(10)
    
    sCadImp = sCadImp & "AGENCIA: " & R!cAgencia & Chr(10)
    sCadImp = sCadImp & String(127, "-") & Chr(10)
    sCadImp = sCadImp & "    FECHA/HORA      USER      NUMERO           ENTIDAD                    MONTO    FECHA     FECHA     AG    Tipo" & Chr(10)
    sCadImp = sCadImp & " REGISTRO DOCUMENTO         DOCUMENTO      DE ORIGEN DEL DOC                      DE REFER  VALORIZAC    " & Chr(10)
    sCadImp = sCadImp & String(127, "-") & Chr(10)
    
    cEstadoTemp = R!cestadocheque
    cAgenciaTemp = R!cAgencia
    
    Do While Not R.EOF
        
        If Trim(cEstadoTemp) <> R!cestadocheque Then
            
            sCadImp = sCadImp & String(127, "-") & Chr(10)
            sCadImp = sCadImp & "TOTAL    " & Left("AGENCIA" & String(15, " "), 15) & ": " & Space(35) & ImpreFormat(nCantidadA, 4, 0) & Space(2) & ImpreFormat(nTotalA, 10, 2) & Chr(10) & Chr(10)
            
            sCadImp = sCadImp & String(127, "-") & Chr(10)
            sCadImp = sCadImp & "TOTAL    " & Left(UCase(cEstadoTemp) & String(15, " "), 15) & ": " & Space(35) & ImpreFormat(nCantidad, 4, 0) & Space(2) & ImpreFormat(nTotal, 10, 2) & Chr(10) & Chr(10)
            
            cEstadoTemp = R!cestadocheque
            nTotal = R!nMonto
            
            cAgenciaTemp = R!cAgencia
            nTotalA = R!nMonto
            
            nCantidad = 1
            nCantidadA = 1
            
            sCadImp = sCadImp & Chr(10)
            sCadImp = sCadImp & "CHEQUES CON ESTADO: " & UCase(R!cestadocheque) & Chr(10) & Chr(10)
            sCadImp = sCadImp & "AGENCIA: " & R!cAgencia & Chr(10)
            sCadImp = sCadImp & String(127, "-") & Chr(10)
            sCadImp = sCadImp & "    FECHA/HORA      USER      NUMERO           ENTIDAD                    MONTO    FECHA     FECHA     AG" & Chr(10)
            sCadImp = sCadImp & " REGISTRO DOCUMENTO         DOCUMENTO      DE ORIGEN DEL DOC                      DE REFER  VALORIZAC    " & Chr(10)
            sCadImp = sCadImp & String(127, "-") & Chr(10)
                        
        Else
            
        
            nTotal = nTotal + R!nMonto
            nCantidad = nCantidad + 1
            
            If cAgenciaTemp <> R!cAgencia Then
                sCadImp = sCadImp & String(127, "-") & Chr(10)
                sCadImp = sCadImp & "TOTAL    " & Left("AGENCIA" & String(15, " "), 15) & ": " & Space(35) & ImpreFormat(nCantidadA, 4, 0) & Space(2) & ImpreFormat(nTotalA, 10, 2) & Chr(10) & Chr(10)
            
                cAgenciaTemp = R!cAgencia
                nTotalA = R!nMonto
                
                nCantidadA = 1
            
                sCadImp = sCadImp & "AGENCIA: " & R!cAgencia & Chr(10)
                sCadImp = sCadImp & String(127, "-") & Chr(10)
                sCadImp = sCadImp & "    FECHA/HORA      USER      NUMERO           ENTIDAD                    MONTO    FECHA     FECHA     AG" & Chr(10)
                sCadImp = sCadImp & " REGISTRO DOCUMENTO         DOCUMENTO      DE ORIGEN DEL DOC                      DE REFER  VALORIZAC    " & Chr(10)
                sCadImp = sCadImp & String(127, "-") & Chr(10)
            
                nTotalA = R!nMonto
                nCantidadA = 1
            Else
                nTotalA = nTotalA + R!nMonto
                nCantidadA = nCantidadA + 1
            End If
            
        End If
        
        cNroDoc = R!cNroDoc
        cPersona = R!persona
    
        sCadImp = sCadImp & R!FechaHora & " " & R!Usu & " " & cNroDoc & " " & cPersona & " " & ImpreFormat(R!nMonto, 10, 2) & " " & R!dValorizaRef & " " & R!dValorizacion & " " & R!cAgencia & " " & R!prod & Chr(10) ' era Chr(10) y agrego Producto
          
        R.MoveNext
    Loop
    
    sCadImp = sCadImp & String(127, "-") & Chr(10)
    sCadImp = sCadImp & "TOTAL    " & Left("AGENCIA" & String(15, " "), 15) & ": " & Space(35) & ImpreFormat(nCantidadA, 4, 0) & Space(2) & ImpreFormat(nTotalA, 10, 2) & Chr(10) & Chr(10)
    
    sCadImp = sCadImp & String(127, "-") & Chr(10)
    sCadImp = sCadImp & "TOTAL    " & Left(UCase(cEstadoTemp) & String(15, " "), 15) & ": " & Space(35) & ImpreFormat(nCantidad, 4, 0) & Space(2) & ImpreFormat(nTotal, 10, 2) & Chr(10)
    
End If
R.Close
Set R = Nothing
    
RepoChequesCompleto = sCadImp
End Function

Public Function GetOtrasOperaciones(pdIni As Date, pdFin As Date, pgsEmpresa As String, pgsNomAge As String, psAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
        
    Dim lnCorr As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lsCadena As String
    
    Dim lsMonedaAnt As String
    Dim lsCabeceraAnt As String
    
    Dim lnAcum As Currency
    
    Dim lsItem As String * 4
    Dim lsFecha As String * 20
    Dim lsUsu As String * 4
    Dim lsImporte As String * 15
    Dim lsDoc As String * 15
    Dim lsReferencia As String * 29
    Dim lsTipoOpe As String * 24
    
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "' "
    End If
    
    SQL = " Select dbo.FechaHoraMov(M.cMovNro) FechaHora, Right(M.cMovNro,4) Usu, MOPV.nMoneda, MOPV.nMovImporte, MOPV.cNroDoc, Replace(Replace(MOPV.cReferencia,Char(10),' '),Char(13),' ')  Referencia," _
        & " (Select cOpeDesc from OpeTpo OP Where OP.cOpeCod = Left(M.cOpeCod,4) + '00') Cabecera, (Select cOpeDesc from OpeTpo OP Where OP.cOpeCod = M.cOpeCod) TipoOpe" _
        & " From MovOpeVarias MOPV" _
        & " Inner Join Mov M On MOPV.nMOvNro = M.nMovNro" _
        & " Where M.nMOvFlag = 0 And M.cOpeCod like '30%' " & sqlAdd & " " _
        & " And M.cOpeCod Not like '3003%' And Left(M.cMovNro,8) Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(pdFin, gsFormatoMovFecha) & "'" _
        & " Order By MOPV.nMoneda, Cabecera"

    lsCadena = ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
        lnCorr = 0
        While Not rs.EOF
            If lsMonedaAnt <> rs!nMoneda Or lsCabeceraAnt <> rs!Cabecera Then
                If lsCadena <> "" Then
                    RSet lsImporte = Format(lnAcum, "#,##0.00")
                    lsCadena = lsCadena & Encabezado(" ;30;" & lsImporte & ";17; ;74;", lnItem, False)
                    lsCadena = lsCadena & Chr(12)
                    lnAcum = 0
                End If
                lsCadena = lsCadena & CabeceraPagina(Trim(rs!Cabecera), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, rs!nMoneda)
                lsCadena = lsCadena & Encabezado("Item;4; ;6;Tipo.Ope;10; ;10;Importe;17; ;8;Documento;10; ;2;Referencia;15; ;5;Usu.;14; ;5;Fecha;10; ;5;", lnItem)
            End If
                        
            lnCorr = lnCorr + 1
            lnItem = lnItem + 1
            
            lsItem = Format(lnCorr, "0000")
            lsFecha = rs!FechaHora
            lsUsu = rs!Usu
            RSet lsImporte = Format(rs!nMovImporte, "#,##0.00")
            RSet lsDoc = rs!cNroDoc
            lsReferencia = rs!Referencia
            lsTipoOpe = rs!TipoOpe
            lnAcum = lnAcum + rs!nMovImporte
        
            lsCadena = lsCadena & lsItem & Space(2) & lsTipoOpe & Space(2) & lsImporte & Space(2) & lsDoc & Space(2) & lsReferencia & Space(2) & lsUsu & Space(2) & lsFecha & Chr(10)
            
            lsMonedaAnt = rs!nMoneda
            lsCabeceraAnt = rs!Cabecera
            
            If lnItem > 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim(rs!Cabecera), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, rs!nMoneda)
                lsCadena = lsCadena & Encabezado("Item;4; ;6;Tipo.Ope;10; ;10;Importe;17; ;8;Documento;10; ;2;Referencia;15; ;5;Usu.;14; ;5;Fecha;10; ;5;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    
        If lsCadena <> "" Then
            RSet lsImporte = Format(lnAcum, "#,##0.00")
            lsCadena = lsCadena & Encabezado(" ;30;" & lsImporte & ";17; ;74;", lnItem, False)
        End If
    End If
    
    GetOtrasOperaciones = lsCadena
    
End Function


Public Function GetChequesRecibidos(pdIni As Date, pdFin As Date, pgsEmpresa As String, pgsNomAge As String, psAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
        
    Dim lnCorr As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lsCadena As String
    
    Dim lsMonedaAnt As String
    Dim lsCabeceraAnt As String
    
    Dim lnAcum As Currency
    
    Dim lsItem As String * 4
    Dim lsFecha As String * 20
    Dim lsUsu As String * 4
    Dim lsImporte As String * 15
    Dim lsNroDoc As String * 20
    Dim lsCtaIf As String * 20
    Dim lsReferencia As String * 29
    Dim lsPersona As String * 24
    Dim ldVal As String * 20
    Dim lsTipoOpe As String
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "' "
    End If
    
    SQL = " Select dbo.FechaHoraMov(M.cMovNro) FechaHora, Right(M.cMovNro,4) Usu, DR.cNroDoc, (Select cPersNombre FROM Persona PE Where PE.cPersCod = DR.cPersCod) Persona, DR.bPlaza, DR.cIFCta, DR.nMonto, DR.dValorizaRef, DR.dValorizacion, DRE.nEstado, nMoneda From MovDoc MD " _
        & " Inner Join Mov M On M.nMovNro = MD.nMovNro" _
        & " Inner Join DocRec DR On MD.nDocTpo = DR.nTpoDoc" _
        & "         And DR.cNroDoc = MD.cDocNro" _
        & " Inner Join DocRecEst DRE On DR.nTpoDoc = DRE.nTpoDoc" _
        & "         And DR.cNroDoc = DRE.cNroDoc And DR.cPersCod = DRE.cPersCod" _
        & "         And DR.cIFTpo = DRE.cIFTpo And DRE.cMovNro = M.cMovNro" _
        & " Where MD.nDocTpo = 47 And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(pdFin, gsFormatoMovFecha) & "' " & sqlAdd & " Order By nMoneda"
    'Sql = " Select dbo.FechaHoraMov(M.cMovNro) FechaHora, Right(M.cMovNro,4) Usu, MOPV.nMoneda, MOPV.nMovImporte, MOPV.cNroDoc, Replace(Replace(MOPV.cReferencia,Char(10),' '),Char(13),' ')  Referencia," _
        & " (Select cOpeDesc from OpeTpo OP Where OP.cOpeCod = Left(M.cOpeCod,4) + '00') Cabecera, (Select cOpeDesc from OpeTpo OP Where OP.cOpeCod = M.cOpeCod) TipoOpe" _
        & " From MovOpeVarias MOPV" _
        & " Inner Join Mov M On MOPV.nMOvNro = M.nMovNro" _
        & " Where M.nMOvFlag = 0 And M.cOpeCod like '30%' " & sqlAdd & " " _
        & " And M.cOpeCod Not like '3003%' And Left(M.cMovNro,8) Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(pdFin, gsFormatoMovFecha) & "'" _
        & " Order By MOPV.nMoneda, Cabecera"

    lsCadena = ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
        lnCorr = 0
        While Not rs.EOF
            If lsMonedaAnt <> rs!nMoneda Then
                If lsCadena <> "" Then
                    RSet lsImporte = Format(lnAcum, "#,##0.00")
                    lsCadena = lsCadena & Encabezado(" ;30;" & lsImporte & ";17; ;74;", lnItem, False)
                    lsCadena = lsCadena & Chr(12)
                    lnAcum = 0
                End If
                lsCadena = lsCadena & CabeceraPagina("REPORTE DE CHEQUES RECIBIDOS", lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, rs!nMoneda)
                lsCadena = lsCadena & Encabezado("Item;4; ;6;Tipo.Ope;10; ;10;Importe;17; ;8;Documento;10; ;2;Referencia;15; ;5;Usu.;14; ;5;Fecha;10; ;5;", lnItem)
            End If
                        
            lnCorr = lnCorr + 1
            lnItem = lnItem + 1
            
            lsItem = Format(lnCorr, "0000")
            lsFecha = rs!FechaHora
            lsUsu = rs!Usu
            RSet lsImporte = Format(rs!nMonto, "#,##0.00")
            RSet lsNroDoc = rs!cNroDoc
            lsReferencia = rs!Referencia
            lsTipoOpe = rs!TipoOpe
            lnAcum = lnAcum + rs!nMovImporte
        
            lsCadena = lsCadena & lsItem & Space(2) & lsNroDoc & Space(2) & lsImporte & Space(2) & lsCtaIf & Space(2) & lsPersona & Space(2) & lsReferencia & Space(2) & ldVal & Space(2) & lsFecha & Space(2) & lsUsu & Chr(10)
            
            lsMonedaAnt = rs!nMoneda
            
            rs.MoveNext
        Wend
    
        If lsCadena <> "" Then
            RSet lsImporte = Format(lnAcum, "#,##0.00")
            lsCadena = lsCadena & Encabezado(" ;30;" & lsImporte & ";17; ;74;", lnItem, False)
        End If
    End If
    
    GetChequesRecibidos = lsCadena
    
End Function

Public Function GetListaOPSolicitadas(pdIni As Date, pdFin As Date, pgsEmpresa As String, pgsNomAge As String, psAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
        
    Dim lnCorr As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lsCadena As String
    
    Dim lsMonedaAnt As String
    Dim lsCabeceraAnt As String
    
    Dim lnAcum As Currency
    
    Dim lsItem As String * 4
    Dim lsFecha As String * 20
    Dim lsUsu As String * 4
    Dim lsImporte As String * 12
    Dim lsCta As String * 18
    Dim lsReferencia As String * 53
    
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "' "
    End If
    
    SQL = " Select dbo.FechaHoraMov(M.cMovNro) FechaHora, M.nMovNro, Right(M.cMovNro,4) Usu, M.cMovDesc, MC.cCtaCod, MC.nMonto, Substring(MC.cCtaCod,9,1) nMoneda from mov  M" _
        & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
        & " Where M.cOpecod = '200602'" _
        & " And M.nMovFlag = 0 And Left(M.cMovNro,8) Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(pdFin, gsFormatoMovFecha) & "'" & sqlAdd & " Order by Substring(MC.cCtaCod,9,1), M.cMovNro"
        
    lsCadena = ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
        lnCorr = 0
        While Not rs.EOF
            If lsMonedaAnt <> rs!nMoneda Then
                If lsCadena <> "" Then
                    RSet lsImporte = Format(lnAcum, "#,##0.00")
                    lsCadena = lsCadena & Encabezado(" ;21;" & lsImporte & ";17; ;81;", lnItem, False)
                    lsCadena = lsCadena & Chr(12)
                    lnAcum = 0
                End If
                lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE TALONARIOS DE OP SOLICITADOS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, rs!nMoneda)
                lsCadena = lsCadena & Encabezado("Item;4; ;6;Cuenta;10; ;5;Importe;12; ;8;Referencia;15; ;25;Usu.;14; ;5;Fecha;10; ;5;", lnItem)
            End If
                        
            lnCorr = lnCorr + 1
            lnItem = lnItem + 1
            
            lsItem = Format(lnCorr, "0000")
            lsFecha = rs!FechaHora
            lsUsu = rs!Usu
            RSet lsImporte = Format(rs!nMonto, "#,##0.00")
            lsReferencia = rs!cMovDesc
            lsCta = rs!Cctacod
            lnAcum = lnAcum + rs!nMonto
        
            lsCadena = lsCadena & lsItem & Space(2) & lsCta & Space(2) & lsImporte & Space(2) & lsReferencia & Space(2) & lsUsu & Space(2) & lsFecha & Chr(10)
            
            lsMonedaAnt = rs!nMoneda
            
            If lnItem > 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim(rs!Cabecera), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, rs!nMoneda)
                lsCadena = lsCadena & Encabezado("Item;4; ;6;Cuenta;10; ;5;Importe;12; ;8;Referencia;15; ;25;Usu.;14; ;5;Fecha;10; ;5;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    
        If lsCadena <> "" Then
            RSet lsImporte = Format(lnAcum, "#,##0.00")
            lsCadena = lsCadena & Encabezado(" ;21;" & lsImporte & ";17; ;81;", lnItem, False)
        End If
    End If
    
    GetListaOPSolicitadas = lsCadena
    
End Function

Public Function GetListaLlamadasCMACS(pdIni As Date, pdFin As Date, pgsEmpresa As String, pgsNomAge As String, psAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
        
    Dim lnCorr As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lsCadena As String
    
    Dim cNomCaja As String * 43
    Dim cCodCtaTransaccion As String * 12
    Dim cDocumento As String * 10
    Dim cCliente As String * 10
    Dim nTempo As Integer
    
    Dim cTempoAge As String
    Dim cTempoUsu As String
    Dim cTempoMon As String
    Dim copecod As String
    Dim cOpeDesc As String
    
    Dim nTotalOpe1 As Double
    Dim nTotalOpe2 As Double
    Dim nTotalMon1 As Double
    Dim nTotalMon2 As Double
    Dim nTotalUsu1 As Double
    Dim nTotalUsu2 As Double
    Dim ntotalAge1 As Double
    Dim ntotalAge2 As Double
        
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "' "
    End If
        
    SQL = "select substring(m1.cMovNro, 18,2) as cAgencia, A.cAgeDescripcion, right(m1.cMovNro, 4) as cUser, "
    SQL = SQL & " m1.cFecha, m1.cHora, m1.cCodCaja, m1.cNomCaja, m1.cMoneda, m1.nMonto, m1.cCodCtaTransaccion, "
    SQL = SQL & " m1.cDocumento, m1.cCliente, m1.cOpeCod, m1.cOpeDesc, "
    SQL = SQL & " isnull(m2.cMonedaComision, '') as cMonedaComision, isnull(m2.nImporteComision,0) as nImporteComision, "
    SQL = SQL & " m1.nMovNro as nMovCmac, m1.cMovNro as cMovCmac, "
    SQL = SQL & " isnull(m2.nMovNroRef, 0) as nMovOpeVarias, isnull(M2.cMovNro,'') as cMovOpeVarias, "
    SQL = SQL & " isnull(m3.nMovNroRef, 0) as nMovRegu, isnull(M3.cMovNro,'') as cMovRegu "
    SQL = SQL & " From "
    SQL = SQL & " ( "
    SQL = SQL & " select substring(m.cmovnro, 7, 2) + '/' + substring(m.cmovnro, 5, 2) + '/' + substring(m.cmovnro, 1, 4) as cFecha, "
    SQL = SQL & " substring(m.cMovNro, 9, 2) + ':' + substring(m.cMovNro, 11, 2) + ':' + substring(m.cMovNro, 13, 2)  as cHora, "
    SQL = SQL & " mc.cPersCod as cCodCaja, p.cpersnombre as cNomCaja, mc.nmoneda as cMoneda, mc.nMonto, "
    SQL = SQL & " isnull(mc.cCtaIF, '') as cCodCtaTransaccion, "
    SQL = SQL & " isnull(mc.cDocumento, '') as cDocumento, "
    SQL = SQL & " isnull(mc.cCliente, '') as cCliente, "
    SQL = SQL & " m.cOpeCod, cOpeDesc=case "
    SQL = SQL & " when m.copecod='260501' then 'LLAMADA DEPOSITO EFECTIVO CMAC' "
    SQL = SQL & " when m.copecod='260503' then 'LLAMADA RETIRO EFECTIVO CMAC' "
    SQL = SQL & " when m.copecod='260504' then 'LLAMADA RETIRO ORDEN DE PAGO' "
    SQL = SQL & " when m.copecod='107001' then 'LLAMADA PAGO DE CREDITO' "
    SQL = SQL & " end, m.nMovNro , m.cMovNro from mov m inner join movcmac mc on m.nmovnro=mc.nmovnro "
    SQL = SQL & " inner join persona p on mc.cperscod=p.cperscod "
    SQL = SQL & " where substring(m.cmovnro,1,8)='" & Format(pdIni, gsFormatoMovFecha) & "' "
    SQL = SQL & " and m.copecod in(260501, 260503, 260504, 107001) "
    SQL = SQL & sqlAdd
    SQL = SQL & " and m.nmovflag=0 ) m1 "
    SQL = SQL & " Inner Join (select cAgeCod, cAgeDescripcion from agencias) A "
    SQL = SQL & " On substring(m1.cMovNro, 18,2) = A.cAgeCod "
    SQL = SQL & " Left Join "
    SQL = SQL & " ( "
    SQL = SQL & " Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef, MO.nMovImporte as nImporteComision, MO.nMoneda as cMonedaComision "
    SQL = SQL & " from MovRef MR Inner Join MovOpevarias MO "
    SQL = SQL & " on MR.nMovNroRef=MO.nMovNro Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef "
    SQL = SQL & " ) m2 "
    SQL = SQL & " on m1.nMovNro = m2.nMovNro "
    SQL = SQL & " Left Join "
    SQL = SQL & " ( "
    SQL = SQL & " Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef, MC.nMonto as nImporteRegu, substring(MC.cCtaCod, 9,1) as cMonedaRegu "
    SQL = SQL & " from MovRef MR Inner Join MovCap MC "
    SQL = SQL & " on MR.nMovNroRef=MC.nMovNro Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef "
    SQL = SQL & " ) m3 "
    SQL = SQL & " on m1.nMovNro = m3.nMovNro "
    SQL = SQL & " Order By substring(m1.cMovNro, 18,2), " '--Agencia
    SQL = SQL & " right(m1.cMovNro, 4), " ' -- User
    SQL = SQL & " m1.cMoneda, " '--Moneda
    SQL = SQL & " m1.cOpeCod, "
    'sql = sql & " m1.cCodCaja, " '-- Caja Municipal
    SQL = SQL & " m1.cMovNro " '--Hora
    
        
    lsCadena = ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
        
        cTempoAge = rs!cAgencia
        cTempoUsu = rs!cUser
        cTempoMon = rs!cMoneda
        copecod = rs!copecod
        cOpeDesc = rs!cOpeDesc
        
        lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - LLAMADAS CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
        lsCadena = lsCadena & String(136, "=") & Chr(10)
        
        lsCadena = lsCadena & "AGENCIA: " & rs!cAgencia & " - " & rs!cAgeDescripcion & Chr(10) & Chr(10)
        lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
        lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
        lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10)
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        
        lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(1) & "COMIS." & Space(1) & "CMAC    " & "REGULAR" & Space(1) & "COMIS." & Chr(10)
        
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        
        lnItem = lnItem + 5
        lnItem = lnItem + 10
        
        Do While Not rs.EOF
   
            If rs!cAgencia <> cTempoAge Then
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(2) & ImpreFormat(nTotalOpe2, 2, 2) & Chr(10) & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Space(2) & ImpreFormat(nTotalMon2, 2, 2) & Chr(10) & Chr(10)
                'lsCadena = lsCadena & String(136, "-") & Chr(10)
                'lsCadena = lsCadena & ImpreFormat("TOTAL x USUARIO    ", 90) & ImpreFormat(nTotalUsu1, 10, 2) & Space(3) & ImpreFormat(nTotalUsu2, 2, 2) & Chr(10) & Chr(10)
                'lsCadena = lsCadena & String(136, "-") & Chr(10)
                'lsCadena = lsCadena & ImpreFormat("TOTAL x AGENCIA    ", 90) & ImpreFormat(nTotalAge1, 10, 2) & Space(3) & ImpreFormat(nTotalAge2, 2, 2) & Chr(10) & Chr(10)
                nTotalOpe1 = rs!nMonto
                nTotalOpe2 = rs!nImporteComision
                nTotalMon1 = rs!nMonto
                nTotalMon2 = rs!nImporteComision
                nTotalUsu1 = rs!nMonto
                nTotalUsu2 = rs!nImporteComision
                ntotalAge1 = rs!nMonto
                ntotalAge2 = rs!nImporteComision
                nTempo = 1
                cTempoAge = rs!cAgencia
                cTempoUsu = rs!cUser
                cTempoMon = rs!cMoneda
                copecod = rs!copecod
                cOpeDesc = rs!cOpeDesc
                               
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - LLAMADAS CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
                lsCadena = lsCadena & String(136, "=") & Chr(10)
                lnItem = 6
                
                lsCadena = lsCadena & "AGENCIA: " & rs!cAgencia & " - " & rs!cAgeDescripcion & Chr(10) & Chr(10)
                lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
                lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(1) & "COMIS." & Space(1) & "CMAC    " & "REGULAR" & Space(1) & "COMIS." & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                
                lnItem = lnItem + 17
                
            Else
                ntotalAge1 = ntotalAge1 + rs!nMonto
                ntotalAge2 = ntotalAge2 + rs!nImporteComision
                
                If rs!cUser <> cTempoUsu Then
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(2) & ImpreFormat(nTotalOpe2, 2, 2) & Chr(10) & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Space(2) & ImpreFormat(nTotalMon2, 2, 2) & Chr(10) & Chr(10)
                    'lsCadena = lsCadena & String(136, "-") & Chr(10)
                    'lsCadena = lsCadena & ImpreFormat("TOTAL x USUARIO    ", 90) & ImpreFormat(nTotalUsu1, 10, 2) & Space(3) & ImpreFormat(nTotalUsu2, 2, 2) & Chr(10) & Chr(10)
                    
                    
                    nTotalOpe1 = rs!nMonto
                    nTotalOpe2 = rs!nImporteComision
                    nTotalMon1 = rs!nMonto
                    nTotalMon2 = rs!nImporteComision
                    nTotalUsu1 = rs!nMonto
                    nTotalUsu2 = rs!nImporteComision
                    nTempo = 1
                    
                    cTempoUsu = rs!cUser
                    cTempoMon = rs!cMoneda
                    copecod = rs!copecod
                    cOpeDesc = rs!cOpeDesc
                    
                    lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
                    lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                    lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(1) & "COMIS." & Space(1) & "CMAC    " & "REGULAR" & Space(1) & "COMIS." & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    
                    lnItem = lnItem + 15
                    
                Else
                    nTotalUsu1 = nTotalUsu1 + rs!nMonto
                    nTotalUsu2 = nTotalUsu2 + rs!nImporteComision
                
                    If rs!cMoneda <> cTempoMon Then
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(2) & ImpreFormat(nTotalOpe2, 2, 2) & Chr(10) & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Space(2) & ImpreFormat(nTotalMon2, 2, 2) & Chr(10) & Chr(10)
                        
                        nTotalOpe1 = rs!nMonto
                        nTotalOpe2 = rs!nImporteComision
                        nTotalMon1 = rs!nMonto
                        nTotalMon2 = rs!nImporteComision
                        nTempo = 1
                        cTempoMon = rs!cMoneda
                        copecod = rs!copecod
                        cOpeDesc = rs!cOpeDesc
                        
                        lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                        lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(1) & "COMIS." & Space(1) & "CMAC    " & "REGULAR" & Space(1) & "COMIS." & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        
                        lnItem = lnItem + 13
                        
                    Else
                        nTotalMon1 = nTotalMon1 + rs!nMonto
                        nTotalMon2 = nTotalMon2 + rs!nImporteComision
                    
                        If rs!copecod <> copecod Then
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(2) & ImpreFormat(nTotalOpe2, 2, 2) & Chr(10) & Chr(10)
                            
                            nTotalOpe1 = rs!nMonto
                            nTotalOpe2 = rs!nImporteComision
                            nTempo = 1
                            copecod = rs!copecod
                            cOpeDesc = rs!cOpeDesc
                            lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10)
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(1) & "COMIS." & Space(1) & "CMAC    " & "REGULAR" & Space(1) & "COMIS." & Chr(10)
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            
                            lnItem = lnItem + 7
                            
                        Else
                            nTempo = nTempo + 1
                            nTotalOpe1 = nTotalOpe1 + rs!nMonto
                            nTotalOpe2 = nTotalOpe2 + rs!nImporteComision
                        End If
                    End If
                End If
            End If
            
            lnItem = lnItem + 1
            
            cNomCaja = rs!cNomCaja
            cCodCtaTransaccion = rs!cCodCtaTransaccion
            cDocumento = rs!cDocumento
            cCliente = rs!cCliente
            
            lsCadena = lsCadena & ImpreFormat(nTempo, 3, 0) & " " & rs!cHora & " " & cNomCaja & " "
            lsCadena = lsCadena & cCodCtaTransaccion & " " & cDocumento & " " & cCliente & " " & ImpreFormat(rs!nMonto, 10, 2) & " " & IIf(rs!cmonedacomision = "1", "S", IIf(rs!cmonedacomision = "2", "D", " "))
            lsCadena = lsCadena & IIf(rs!nImporteComision = 0, "     ", ImpreFormat(rs!nImporteComision, 2, 2))
            
            
            lsCadena = lsCadena & " " & ImpreFormat(rs!nMovcmac, 7, 0)
            
            If Len(Trim(rs!cMovRegu)) = 0 Then
                lsCadena = lsCadena & "        "
            Else
                lsCadena = lsCadena & " " & ImpreFormat(rs!nMovRegu, 7, 0)
            End If
            If Len(Trim(rs!cMovOpevarias)) = 0 Then
                lsCadena = lsCadena & "        "
            Else
                lsCadena = lsCadena & " " & ImpreFormat(rs!nMovOpeVarias, 7, 0)
            End If
            
            lsCadena = lsCadena & Chr(10)
            
            If lnItem >= 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - LLAMADAS CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
                lsCadena = lsCadena & String(136, "=") & Chr(10)
                lnItem = 6
            End If
            
            rs.MoveNext
        Loop
                
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(5) & ImpreFormat(nTotalOpe2, 2, 2) & Chr(10) & Chr(10)
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Space(5) & ImpreFormat(nTotalMon2, 2, 2) & Chr(10) & Chr(10)
        'lsCadena = lsCadena & String(136, "-") & Chr(10)
        'lsCadena = lsCadena & ImpreFormat("TOTAL x USUARIO    ", 90) & ImpreFormat(nTotalUsu1, 10, 2) & Space(3) & ImpreFormat(nTotalUsu2, 2, 2) & Chr(10) & Chr(10)
        'lsCadena = lsCadena & String(136, "-") & Chr(10)
        'lsCadena = lsCadena & ImpreFormat("TOTAL x AGENCIA    ", 90) & ImpreFormat(nTotalAge1, 10, 2) & Space(3) & ImpreFormat(nTotalAge2, 2, 2) & Chr(10) & Chr(10)
        
    End If
    
    GetListaLlamadasCMACS = lsCadena
    
End Function

Public Function GetListaRecepcionCMACS(pdIni As Date, pdFin As Date, pgsEmpresa As String, pgsNomAge As String, psAge As String, pgdFecSis As Date) As String
    Dim SQL As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
        
    Dim lnCorr As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lsCadena As String
    
    Dim cNomCaja As String * 43
    Dim cCodCtaTransaccion As String * 12
    Dim cDocumento As String * 10
    Dim cCliente As String * 10
    Dim nTempo As Integer
    
    Dim cTempoAge As String
    Dim cTempoUsu As String
    Dim cTempoMon As String
    Dim copecod As String
    Dim cOpeDesc As String
    
    Dim nTotalOpe1 As Double
    Dim nTotalMon1 As Double
    Dim nTotalUsu1 As Double
    Dim ntotalAge1 As Double
        
    Dim sqlAdd As String
    
    If psAge = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And Substring(M.cMovNro,18,2) = '" & psAge & "' "
    End If
        
    SQL = "select substring(m1.cMovNro, 18,2) as cAgencia, A.cAgeDescripcion, right(m1.cMovNro, 4) as cUser, "
    SQL = SQL & " m1.cFecha, m1.cHora, m1.cCodCaja, m1.cNomCaja, m1.cMoneda, m1.nMonto, m1.cCodCtaTransaccion, "
    SQL = SQL & " m1.cDocumento, m1.cCliente, m1.cOpeCod, m1.cOpeDesc, M1.nMovNro "
    SQL = SQL & " From "
    SQL = SQL & " ( "
    
    SQL = SQL & " select substring(m.cmovnro, 7, 2) + '/' + substring(m.cmovnro, 5, 2) + '/' + substring(m.cmovnro, 1, 4) as cFecha, "
    SQL = SQL & " substring(m.cMovNro, 9, 2) + ':' + substring(m.cMovNro, 11, 2) + ':' + substring(m.cMovNro, 13, 2)  as cHora, "
    SQL = SQL & " mc.cPersCod as cCodCaja, p.cpersnombre as cNomCaja, mc.nmoneda as cMoneda, mc.nMonto, "
    SQL = SQL & " isnull(mc.cCtaIF, '') as cCodCtaTransaccion,  isnull(mc.cDocumento, '') as cDocumento, "
    SQL = SQL & " isnull(mc.cCliente, '') as cCliente,  m.cOpeCod, "
    SQL = SQL & " ot.cOpeDesc, "
    SQL = SQL & " m.nMovNro , m.cMovNro "
    SQL = SQL & " from mov m inner join movcmac mc on m.nmovnro=mc.nmovnro "
    SQL = SQL & " inner join persona p on mc.cperscod=p.cperscod "
    SQL = SQL & " inner join opetpo ot on m.copecod=ot.copecod "
    SQL = SQL & " where substring(m.cmovnro,1,8)='" & Format(pdIni, gsFormatoMovFecha) & "' "
    'sql = sql & " and m.copecod in('260101', '260102','260103','260104','260201','260301','260302', '126100', '126200', '106001', '136101') "
    
    SQL = SQL & " and m.copecod in('100205','100305','100405','107001','126100','126101','126102','126200','126202','136301','260101','260102','260103','260104','260201','260301','260302','260501','260503','260504','100406','106001', '136101') "
    
    SQL = SQL & sqlAdd
    SQL = SQL & " and m.nmovflag=0 ) m1 "
    
    SQL = SQL & " Inner Join (select cAgeCod, cAgeDescripcion from agencias) A  On substring(m1.cMovNro, 18,2) = A.cAgeCod "
    
    SQL = SQL & " Order By substring(m1.cMovNro, 18,2), " '--Agencia
    SQL = SQL & " right(m1.cMovNro, 4), " ' -- User
    SQL = SQL & " m1.cMoneda, " '--Moneda
    SQL = SQL & " m1.cOpeCod, "
    'sql = sql & " m1.cCodCaja, " '-- Caja Municipal
    SQL = SQL & " m1.cMovNro " '--Hora
        
    lsCadena = ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(SQL)
    
    If Not (rs.EOF And rs.BOF) Then
        
        cTempoAge = rs!cAgencia
        cTempoUsu = rs!cUser
        cTempoMon = rs!cMoneda
        copecod = rs!copecod
        cOpeDesc = rs!cOpeDesc
        
        lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - RECEPCIONES CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
        lsCadena = lsCadena & String(136, "=") & Chr(10)
        
        lsCadena = lsCadena & "AGENCIA: " & rs!cAgencia & " - " & rs!cAgeDescripcion & Chr(10) & Chr(10)
        lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
        lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
        lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10)
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        
        lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(2) & "OPERAC" & Chr(10)
        
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        
        lnItem = lnItem + 5
        lnItem = lnItem + 10
        
        Do While Not rs.EOF
   
            If rs!cAgencia <> cTempoAge Then
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Chr(10) & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Chr(10) & Chr(10)
                'lsCadena = lsCadena & String(136, "-") & Chr(10)
                'lsCadena = lsCadena & ImpreFormat("TOTAL x USUARIO    ", 90) & ImpreFormat(nTotalUsu1, 10, 2) & Space(3) & ImpreFormat(nTotalUsu2, 2, 2) & Chr(10) & Chr(10)
                'lsCadena = lsCadena & String(136, "-") & Chr(10)
                'lsCadena = lsCadena & ImpreFormat("TOTAL x AGENCIA    ", 90) & ImpreFormat(nTotalAge1, 10, 2) & Space(3) & ImpreFormat(nTotalAge2, 2, 2) & Chr(10) & Chr(10)
                nTotalOpe1 = rs!nMonto
                nTotalMon1 = rs!nMonto
                nTotalUsu1 = rs!nMonto
                ntotalAge1 = rs!nMonto
                nTempo = 1
                cTempoAge = rs!cAgencia
                cTempoUsu = rs!cUser
                cTempoMon = rs!cMoneda
                copecod = rs!copecod
                cOpeDesc = rs!cOpeDesc
                               
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - RECEPCIONES CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
                lsCadena = lsCadena & String(136, "=") & Chr(10)
                lnItem = 6
                
                lsCadena = lsCadena & "AGENCIA: " & rs!cAgencia & " - " & rs!cAgeDescripcion & Chr(10) & Chr(10)
                lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
                lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(2) & "OPERAC" & Chr(10)
                lsCadena = lsCadena & String(136, "-") & Chr(10)
                
                lnItem = lnItem + 17
                
            Else
                ntotalAge1 = ntotalAge1 + rs!nMonto
                                
                If rs!cUser <> cTempoUsu Then
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Chr(10) & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Chr(10) & Chr(10)
                    'lsCadena = lsCadena & String(136, "-") & Chr(10)
                    'lsCadena = lsCadena & ImpreFormat("TOTAL x USUARIO    ", 90) & ImpreFormat(nTotalUsu1, 10, 2) & Space(3) & ImpreFormat(nTotalUsu2, 2, 2) & Chr(10) & Chr(10)
                    
                    
                    nTotalOpe1 = rs!nMonto
                    nTotalMon1 = rs!nMonto
                    nTotalUsu1 = rs!nMonto
                    nTempo = 1
                    
                    cTempoUsu = rs!cUser
                    cTempoMon = rs!cMoneda
                    copecod = rs!copecod
                    cOpeDesc = rs!cOpeDesc
                    
                    lsCadena = lsCadena & "USUARIO: " & rs!cUser & Chr(10) & Chr(10)
                    lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                    lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(2) & "OPERAC" & Chr(10)
                    lsCadena = lsCadena & String(136, "-") & Chr(10)
                    
                    lnItem = lnItem + 15
                    
                Else
                    nTotalUsu1 = nTotalUsu1 + rs!nMonto
                                    
                    If rs!cMoneda <> cTempoMon Then
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Space(2) & Chr(10) & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Space(2) & Chr(10) & Chr(10)
                        
                        nTotalOpe1 = rs!nMonto
                        nTotalMon1 = rs!nMonto
                        nTempo = 1
                        cTempoMon = rs!cMoneda
                        copecod = rs!copecod
                        cOpeDesc = rs!cOpeDesc
                        
                        lsCadena = lsCadena & "MONEDA: " & IIf(rs!cMoneda = "1", "Soles", "Dolares") & Chr(10) & Chr(10)
                        lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10) & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(2) & "OPERAC" & Chr(10)
                        lsCadena = lsCadena & String(136, "-") & Chr(10)
                        
                        lnItem = lnItem + 13
                        
                    Else
                        nTotalMon1 = nTotalMon1 + rs!nMonto
                        
                        If rs!copecod <> copecod Then
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Chr(10) & Chr(10)
                            
                            nTotalOpe1 = rs!nMonto
                            nTempo = 1
                            copecod = rs!copecod
                            cOpeDesc = rs!cOpeDesc
                            lsCadena = lsCadena & "OPERACION: " & rs!copecod & " - " & rs!cOpeDesc & Chr(10)
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            lsCadena = lsCadena & "ITEM HORA   " & Space(1) & "CAJA  MUNICIPAL  DE  REFERENCIA    " & Space(9) & "CUENTA      " & Space(1) & "DOCUMENTO " & Space(1) & "CLIENTE   " & "  M O N T O  " & Space(1) & Space(2) & "OPERAC" & Chr(10)
                            lsCadena = lsCadena & String(136, "-") & Chr(10)
                            
                            lnItem = lnItem + 7
                            
                        Else
                            nTempo = nTempo + 1
                            nTotalOpe1 = nTotalOpe1 + rs!nMonto
                        End If
                    End If
                End If
            End If
            
            lnItem = lnItem + 1
            
            cNomCaja = rs!cNomCaja
            cCodCtaTransaccion = rs!cCodCtaTransaccion
            cDocumento = rs!cDocumento
            cCliente = rs!cCliente
            
            lsCadena = lsCadena & ImpreFormat(nTempo, 3, 0) & " " & rs!cHora & " " & cNomCaja & " "
            lsCadena = lsCadena & cCodCtaTransaccion & " " & cDocumento & " " & cCliente & " " & ImpreFormat(rs!nMonto, 10, 2) & " "
            
            lsCadena = lsCadena & " " & ImpreFormat(rs!nMovNro, 7, 0)
            
            lsCadena = lsCadena & Chr(10)
            
            If lnItem >= 54 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & CabeceraPagina(Trim("LISTADO DE OPERACIONES - RECEPCIONES CON CMACS"), lnPagina, lnItem, gsNomAge, pgsEmpresa, pgdFecSis, , False)
                lsCadena = lsCadena & String(136, "=") & Chr(10)
                lnItem = 6
            End If
            
            rs.MoveNext
        Loop
                
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        lsCadena = lsCadena & ImpreFormat("TOTAL x OPERACION  ", 90) & ImpreFormat(nTotalOpe1, 10, 2) & Chr(10) & Chr(10)
        lsCadena = lsCadena & String(136, "-") & Chr(10)
        lsCadena = lsCadena & ImpreFormat("TOTAL x MONEDA     ", 90) & ImpreFormat(nTotalMon1, 10, 2) & Chr(10) & Chr(10)
        
    End If
    
    GetListaRecepcionCMACS = lsCadena
    
End Function




