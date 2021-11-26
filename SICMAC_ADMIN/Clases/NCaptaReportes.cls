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
Private I As Integer
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

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3A835D5E0029
    Private mlClassDebugID As Long
#End If

'##ModelId=3A7B7F2C0285
Public theNImprimir As NImprimir

'##ModelId=3A7A69160086
'Public theNCapMantenimiento As NCapMantenimiento

'Obtiene la plantilla de las ordenes de pago asigna valores a sus campos
'##ModelId=3A7B819F024B

'##ModelId=3A835D5E021E
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
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
Public Function Reporte(psRep As String, psFecIni As String, psFecFin As String, psNomAge As String, psEmpresa As String, pdFecSis As Date, psCodAge As String, Optional psUser As String = "XXX") As String
    Dim lsCadena As String
    Dim lsRep As String
    Dim lsRepSub As String
    lsNomAge = psNomAge
    lsEmpresa = psEmpresa
    ldFecSis = pdFecSis
    ldFecRepo = CDate(psFecIni)
    lsAgeCod = psCodAge
    lnRango = 54
    
    lsRep = psRep
    While lsRep <> ""
        lsRepSub = Left(lsRep, 3)
        lsRep = Mid(lsRep, 5)
        If lsRepSub = "D01" Then
            lsCadena = lsCadena & LlenaEDAC(1, psFecIni, psFecFin)
            lsCadena = lsCadena & LlenaEDAC(2, psFecIni, psFecFin)
        ElseIf lsRepSub = "D02" Then
            lsCadena = lsCadena & LlenaEDPF(1, psFecIni, psFecFin)
            lsCadena = lsCadena & LlenaEDPF(2, psFecIni, psFecFin)
        ElseIf lsRepSub = "D03" Then
            lsCadena = lsCadena & LlenaEDCTS(1, psFecIni, psFecFin)
            lsCadena = lsCadena & LlenaEDCTS(2, psFecIni, psFecFin)
        ElseIf lsRepSub = "D04" Then
            lsCadena = lsCadena & ReporteTrasTotSM("MOVIMIENTOS DE AHORRO", False, , Format(CDate(psFecIni), "yyyymmdd"))
        ElseIf lsRepSub = "D05" Then
            lsCadena = lsCadena & SdosTipCta(lsAgeCod)
        ElseIf lsRepSub = "D06" Then
            lsCadena = lsCadena & EstratDep(lsAgeCod)
        ElseIf lsRepSub = "M01" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "0", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "0", "2", psFecIni)
        ElseIf lsRepSub = "M02" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "1", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "1", "2", psFecIni)
        ElseIf lsRepSub = "M03" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "2", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "2", "2", psFecIni)
        ElseIf lsRepSub = "M04" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "3", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapAhorros, "3", "2", psFecIni)
        ElseIf lsRepSub = "M05" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "0", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "0", "2", psFecIni)
        ElseIf lsRepSub = "M06" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "1", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "1", "2", psFecIni)
        ElseIf lsRepSub = "M07" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "2", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "2", "2", psFecIni)
        ElseIf lsRepSub = "M08" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "3", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapPlazoFijo, "3", "2", psFecIni)
        ElseIf lsRepSub = "M09" Then
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, "0", "1", psFecIni)
            lsCadena = lsCadena & GeneraSaldos(Producto.gCapCTS, "0", "2", psFecIni)
        ElseIf lsRepSub = "M10" Then
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "0", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "0", "2", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "1", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "1", "2", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "2", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "2", "2", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "3", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "3", "2", psFecIni, 0)
        
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "0", "1", psFecIni, 1)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "0", "2", psFecIni, 1)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "1", "1", psFecIni, 1)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "1", "2", psFecIni, 1)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "2", "1", psFecIni, 1)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapAhorros, "2", "2", psFecIni, 1)
        
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "0", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "0", "2", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "1", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "1", "2", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "2", "1", psFecIni, 0)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapPlazoFijo, "2", "2", psFecIni, 0)
        
           lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, "0", "1", psFecIni, 0, gsInstCmac)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, "0", "2", psFecIni, 0, "")
           lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, "0", "1", psFecIni, 0, gsInstCmac)
           lsCadena = lsCadena & GenSdosFM(Producto.gCapCTS, "0", "2", psFecIni, 0, "")
        ElseIf lsRepSub = "M11" Then
            lsCadena = lsCadena & RepEstratPlazoFijo
        ElseIf lsRepSub = "V01" Then
            lsCadena = lsCadena & RepCtaIna
        ElseIf lsRepSub = "V02" Then
            lsCadena = lsCadena & ConsolidadoInactivas(CDate(psFecIni), CDate(psFecFin))
        ElseIf lsRepSub = "V03" Then
            lsCadena = lsCadena & CuentasInactivas("1", CDate(psFecIni), CDate(psFecFin))
            lsCadena = lsCadena & CuentasInactivas("2", CDate(psFecIni), CDate(psFecFin))
        ElseIf lsRepSub = "V04" Then
            lsCadena = lsCadena & RepCtaApe(CDate(psFecIni), CDate(psFecFin))
        ElseIf lsRepSub = "V05" Then
            lsCadena = lsCadena & RepCtaCan(CDate(psFecIni), CDate(psFecFin))
        End If
    Wend
    Reporte = lsCadena
End Function

Private Function ReporteTrasTotSM(psTitulo As String, Optional pbFinPg As Boolean = True, Optional psCodUser As String = "XXX", Optional psFecha As String = "") As String
    'On Error GoTo ERROR
    Dim lsCadena As String
    Dim lsAdd As String
    lsCadena = ""
    
    lsAdd = IIf(psFecha = "", "", " Del " & psFecha)
    ReporteMovAC lsCadena, "Detalle de Operaciones Ahorro - Soles" & lsAdd, 0, 0, "1", psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPF lsCadena, "Detalle de Operaciones Plazo Fijo - Soles" & lsAdd, 0, 0, "1", psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCTS lsCadena, "Detalle de Operaciones CTS - Soles" & lsAdd, 0, 0, "1", psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovAC lsCadena, "Detalle de Operaciones Ahorro - Dolares" & lsAdd, 0, 0, "2", psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovPF lsCadena, "Detalle de Operaciones Plazo Fijo - Dolares" & lsAdd, 0, 0, "2", psCodUser, psFecha
    PonerSalto lsCadena
    ReporteMovCTS lsCadena, "Detalle de Operaciones CTS - Dolares", 0, 0, "2" & lsAdd, psCodUser, psFecha
        
    If pbFinPg Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina

    'QuitarSalto lsCadena

    ReporteTrasTotSM = lsCadena
    
    Exit Function
ERROR:
    MsgBox "Error Realizando el Reporte de Movimientos de Cuentas: " & Err.Description, vbCritical, "Aviso"
End Function

'************************************************************************
'Ahorros
'********************************************************************************************
Public Sub ReporteMovAC(psCadenaRep As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", _
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
    
    Dim I As Long
    Dim lsTEM As String
    Dim lsAdd As String
    lnTotRet = 0
    lnTotDep = 0
    
    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))
    
    If psCodUser = "XXX" Then
        lsCabecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    Else
        lsCabecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
    End If
    
    lsTEM = psCadenaRep
    
    psCadenaRep = ""
    
    pnPagina = pnPagina - 1
    
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas de Ahorro en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, Trim(Str(CaptacOperacion.gAhoApeEfec)))     'OK
    Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas de Ahorro con Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, Trim(Str(CaptacOperacion.gAhoApeChq)))    'OK
    'Call ReporteMovACDet(psCadenaRep, "Apertura de Cuentas de Ahorro con NA" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, Trim(Str(CaptacOperacion.gAhoApeEfec)))  'OK
    'Call MovimientoEfectivoAC(psCadenaRep, "Movimientos en Efectivo de Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoDepEfec & "','" & CaptacOperacion.gAhoRetEfec & "','" & CaptacOperacion.gAhoRetOP & "','" & CaptacOperacion.gAhoRetOPCanje & "','" & CaptacOperacion.gAhoRetNC & "','" & CaptacOperacion.gAhoDepNA & "','" & CaptacOperacion.gAhoOPEmision & "','" & CaptacOperacion.gAhoRetOPCert & "','" & CaptacOperacion.gAhoRetOPCertCanje & "','" & CaptacOperacion.gAhoOPCertificacion & "','" & CaptacOperacion.gAhoRetAnulChq)
    'Call ReporteMovACDet(psCadenaRep, "Cancelación de Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoExtCancAct & "','" & CaptacOperacion.gAhoCancNCAct & "','" & CaptacOperacion.gAhoCancNCInact)
    'Call ReporteMovACDet(psCadenaRep, "Depositos con Cheque en Cuentas de Ahorro" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gAhoDepChq)    'OK
    
    If lnTotRet <> 0 Or lnTotDep <> 0 Then
        psCadenaRep = lsCabecera + psCadenaRep
        QuitarSalto psCadenaRep
        psCadenaRep = psCadenaRep + Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
        psCadenaRep = psCadenaRep + Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - Abs(lnTotRet), "###,##0.00")) & ";57; ;46;", pnItem)
        psCadenaRep = psCadenaRep & oImpresora.gPrnSaltoPagina
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

Public Function ReporteMovACDet(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim I As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnTDepT As Double
    Dim lnTRetT As Double
    Dim lnValor As Double
    Dim lbBan As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lbBan = False
    
    lnTDep = 0
    lnTRet = 0
    lnTDepT = 0
    lnTRetT = 0
    
    oCon.AbreConexion
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
     lbBan = True
     psCadena = psCadena & CabeceraPagina(psTitulo & "-Efectivo-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
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
         
         If IsNull(RegTrandiaria!nsalddispac) Then
             lnValor = 0
         Else
             lnValor = RegTrandiaria!nsalddispac
         End If
         
         psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                             & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                             & Space(19 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                             & Space(10 - Len(Right(RegTrandiaria!cMovNro, 4))) & Right(RegTrandiaria!cMovNro, 4) _
                             & Space(17 - Len(Mid(RegTrandiaria!cMovNro, 9, 6)) - 2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                             & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                             & Space(11 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                              
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
     If pnItem < lnRango And lbBan Then
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
     End If
    
     psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
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
         
         If IsNull(RegTrandiaria!nsalddispac) Then
             lnValor = 0
         Else
             lnValor = RegTrandiaria!nsalddispac
         End If
         
         psCadena = psCadena & Space(20 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                            & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                            & Space(19 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                            & Space(10 - Len(Right(RegTrandiaria!cMovNro, 4))) & Right(RegTrandiaria!cMovNro, 4) _
                            & Space(17 - Len(Left(RegTrandiaria!cMovNro, 6)) - 2) & Mid(RegTrandiaria!cMovNro, 9, 2) & ":" & Mid(RegTrandiaria!cMovNro, 11, 2) & ":" & Mid(RegTrandiaria!cMovNro, 13, 2) _
                            & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                            & Space(11 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                            
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
        
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
End Function

Public Function MovimientoEfectivoAC(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
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
    
    Dim I As Long
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lnTRet = 0
    lnTDep = 0
    
    oCon.AbreConexion
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,1,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,1,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
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
                        
'            If RegTrandiaria!cCodOpe = CaptacOperacion.gAhoDepEfec Or RegTrandiaria!cCodOpe = CaptacOperacion.gAhoDepNA Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = RegTrandiaria!nMonTran
'            Else
'                lsDep = "0.00"
'                lsRet = Format(-1 * RegTrandiaria!nMonTran, "###,##0.00")
'                lnTRet = lnTRet + (RegTrandiaria!nMonTran * -1)
'                lnOpe = RegTrandiaria!nMonTran * -1
'            End If
            
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
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,1,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,1,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.BOF And RegTrandiaria.EOF) Then
        If pnItem < lnRango And lbBan Then
            psCadena = psCadena & oImpresora.gPrnSaltoPagina
        End If
        
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sal.Cnt.Ant;20;Deposito;20;Retiro;20;Usu.;6;Hora;12;TInt;8;", pnItem)
        
        While Not RegTrandiaria.EOF
            
            If IsNull(RegTrandiaria!cNumDoc) Then
                lsDoc = " "
            Else
                lsDoc = RegTrandiaria!cNumDoc
            End If
            
            
'            If RegTrandiaria!cCodOpe = CaptacOperacion.gAhoDepEfec Or RegTrandiaria!cCodOpe = CaptacOperacion.gAhoDepNA Then
'                lsDep = Format(RegTrandiaria!nMonTran, "###,##0.00")
'                lsRet = "0.00"
'                lnTDep = lnTDep + RegTrandiaria!nMonTran
'                lnOpe = Abs(RegTrandiaria!nMonTran)
'            Else
'                lsDep = "0.00"
'                lsRet = Format(Abs(RegTrandiaria!nMonTran) * -1, "###,##0.00")
'                lnTRet = lnTRet + Abs(RegTrandiaria!nMonTran) * -1
'                lnOpe = Abs(RegTrandiaria!nMonTran)
'                lnOpe = lnOpe * -1
'            End If
            
            
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
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
        
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
    
End Function

Private Sub PonerSalto(rtf As String)
    If rtf = "" Then Exit Sub
    If Right(rtf, 1) <> oImpresora.gPrnSaltoPagina Then
        rtf = rtf & oImpresora.gPrnSaltoPagina
    End If
End Sub

Private Sub QuitarSalto(rtf As String)
    If Right(rtf, 1) = oImpresora.gPrnSaltoPagina Then
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
'                               & Space(3 - Len(lsOP)) & lsOP & oImpresora.gPrnSaltoLinea
'            pnItem = pnItem + 1
'
'            If pnItem = lnRango Then
'                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
'        psCadena = psCadena & oImpresora.gPrnSaltoPagina
'    End If
'End Function

'Movimiento que se realiza con cuentas de esta agencia pero
'por otro servidor
Private Function MovimientoExternoAC(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String, Optional pbEsTramite As Boolean = False)
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
    
    Dim I As Long
    
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
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(M.cMovNro,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And bOrdPag = 0 And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
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
            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
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
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
              & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacAhorros CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And bOrdPag = 1 And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
              & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
    End If
    
    Set RegTrandiaria = oCon.CargaRecordSet(Sql1)
    
    If Not (RegTrandiaria.EOF And RegTrandiaria.BOF) Then
        If pnItem < lnRango And lbBan Then
            psCadena = psCadena & oImpresora.gPrnSaltoPagina
        End If
        
        lbBan = False
        
        psCadena = psCadena & CabeceraPagina(psTitulo & "-Orden de Pago-", pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        psCadena = psCadena & Encabezado("Nro Cuenta;14;Nro.Doc.;18;Sald.Cnt.Ant.;20;Deposito;20;Retiro;20;UsuR.;6;Hora;12;TInt;8;", pnItem)
    
        lsGru = ""
        
        While Not RegTrandiaria.EOF
            
            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;26;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
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
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntAC, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntAC, "###,##0.00") & oImpresora.gPrnSaltoLinea
                               
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
End Function

'************************************************************************
'PLAZO FIJO
'********************************************************************************************
Private Sub ReporteMovPF(psCadenaRep As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'
'    Dim SQL1 As String
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lnDep As Double
'    Dim lnRet As Double
'
'    Dim lnTotDep As Double
'    Dim lnTotRet As Double
'    Dim lsCebecera As String
'
'    Dim lsTEM As String
'    Dim i As Long
'    Dim lsAdd As String
'    lnTotRet = 0
'    lnTotDep = 0
'
'    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))
'    If psCodUser = "XXX" Then
'        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
'    Else
'        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
'    End If
'
'    pnPagina = pnPagina - 1
'
'    lsTEM = psCadenaRep
'
'    psCadenaRep = ""
'
'    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas PlazoFijo-Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFApeEfec)
'    Call ReporteAperturasPF(psCadenaRep, "Movimientos en Efectivo de Cuentas de Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFRetInt & "','" & CaptacOperacion.gPFRetIntAboAho)
'    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas Plazo Fijo-Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFApeChq)
'    Call ReporteAperturasPF(psCadenaRep, "Apertura de Cuentas Plazo Fijo-NA" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFApeNA)
'    Call ReporteAperturasPF(psCadenaRep, "Cancelación de Cuentas Plazo Fijo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFCancelacion)
'
'    If lnTotDep <> 0 Or lnTotRet <> 0 Then
'        psCadenaRep = lsCebecera + psCadenaRep
'        QuitarSalto psCadenaRep
'        psCadenaRep = psCadenaRep + Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
'        psCadenaRep = psCadenaRep + Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
'        psCadenaRep = psCadenaRep & oImpresora.gPrnSaltoPagina
'    End If
'
'    lnTotRet = 0
'    lnTotDep = 0
'
'    Call MovimientoExternosPF(psCadenaRep, "Cuentas de -Plazo Fijo- Movidas por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFRetInt & "','" & CaptacOperacion.gPFRetIntAboAho)
'    Call MovimientoExternosPF(psCadenaRep, "Cuentas Canceladas de -Plazo Fijo- por otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFCancelacion)
'    Call MovimientoExternosPF(psCadenaRep, "Cuentas de -Plazo Fijo- Movidas en otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, 0, 0, psMoneda, psCodUser, psFecha, CaptacOperacion.gPFRetInt & "','" & CaptacOperacion.gPFRetIntAboAho, True)
'
'    psCadenaRep = lsTEM + psCadenaRep
'    lsTEM = ""
End Sub

Private Sub ReporteAperturasPF(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As ADODB.Recordset
    Set RegTrandiaria = New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim I As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispPF, nTasaInteres nTasaIntPF, nPlazo Plazo, nSaldoContable nSaldCnt" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispPF, nTasaInteres nTasaIntPF, nSaldoContable nSaldCnt" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
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
                             & Space(10 - Len(Str(RegTrandiaria!Plazo))) & Str(RegTrandiaria!Plazo) & oImpresora.gPrnSaltoLinea
         pnItem = pnItem + 1
                 
         If pnItem = lnRango Then
             psCadena = psCadena & oImpresora.gPrnSaltoPagina
             psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
             psCadena = psCadena & Encabezado("Nro Cuenta;15;Nro.Doc.;20;Monto;20;Usu.;9;Hora;15;Sald.Disp.;20;TInt;9;D.Plazo;10", pnItem)
         End If
         
         RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;15;" & Trim(Format(lnTDep, "###,##0.00")) & ";40; ;63;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
End Sub

'Operacioens hechas por otras agencias en este servidor
Private Function MovimientoExternosPF(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String = "", Optional pbEsTramite As Boolean = False)
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

    
    Dim I As Long
    
    lnTRet = 0
    lnTDep = 0
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If pbEsTramite Then
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispPF, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(M.cMovNro,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispPF, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge, nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge , nPlazo Plazo" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacPlazoFijo CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
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
            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15; ;37;" & Format(lnTRetAge * -1, "###,##0.00") & ";20; ;46;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If
            
            
'
            If Len(RegTrandiaria!cCodCta) = 18 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
                    psCadena = psCadena & Encabezado("Agencia: " & RegTrandiaria!cAgeDescripcion & ";32;", pnItem)
                    lbBan = True
                End If
            ElseIf Len(RegTrandiaria!cCodCta) = 8 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
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
                               & Space(17 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
                               & Space(8 - Len(Format(RegTrandiaria!nTasaIntPF, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntPF, "###,##0.00") _
                               & Space(11 - Len(Str(RegTrandiaria!nPlazo))) & Str(RegTrandiaria!nPlazo) & oImpresora.gPrnSaltoLinea
            
            'Str(RegTrandiaria!nPlazo)
            pnItem = pnItem + 1
                    
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
        
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
End Function


'************************************************************************
'CTS
'********************************************************************************************
Private Sub ReporteMovCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "")
'
'    Dim SQL1 As String
'    Dim lsDoc As String
'    Dim lsCodCta As String
'    Dim lnDep As Double
'    Dim lnRet As Double
'
'    Dim lnTotDep As Double
'    Dim lnTotRet As Double
'    Dim lsCebecera As String
'
'    Dim lsTEM As String
'    Dim i As Long
'    Dim lsAdd As String
'
'    lsAdd = IIf(psFecha = "", "", " Del " & Right(psFecha, 2) & "/" & Mid(psFecha, 5, 2) & "/" & Left(psFecha, 4))
'
'    lnTotRet = 0
'    lnTotDep = 0
'
'    If psCodUser = "XXX" Then
'        lsCebecera = CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
'    Else
'        lsCebecera = CabeceraPagina(psTitulo & " " & psCodUser, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
'    End If
'
'    pnPagina = pnPagina - 1
'    lsTEM = psCadena
'
'    psCadena = ""
'
'    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS en Efectivo" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSExtApeEfec)
'    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS con Cheque" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSExtApeChq)
'    Call ReporteAperturasCTS(psCadena, "Apertura de Cuentas CTS con NC" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSExtApeNA)
'    Call ReporteAperturasCTS(psCadena, "Movimientos en Efectivo de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSDepEfec & "','" & CaptacOperacion.gCTSRetEfec & "','" & CaptacOperacion.gCTSExtDepNA & "','" & CaptacOperacion.gCTSRetNC & "','" & CaptacOperacion.gCTSRetEfec)
'    Call ReporteAperturasCTS(psCadena, "Cancelación de Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSCancelacion)
'    Call ReporteAperturasCTS(psCadena, "Depositos con Cheque en Cuentas CTS" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha, CaptacOperacion.gCTSDepChq)
'
'    If lnTotRet <> 0 Or lnTotDep <> 0 Then
'        psCadena = lsCebecera + psCadena
'        QuitarSalto psCadena
'        psCadena = psCadena & Encabezado("Resumen Total;15;" & Trim(Format(lnTotDep, "###,##0.00")) & ";57;" & Trim(Format(lnTotRet, "###,##0.00")) & ";20; ;26;", pnItem)
'        psCadena = psCadena & Encabezado("Monto Total;15;" & Trim(Format(lnTotDep - lnTotRet, "###,##0.00")) & ";57; ;46;", pnItem)
'        psCadena = psCadena & oImpresora.gPrnSaltoPagina
'    End If
'
'    'Call MovimientoExternosCTS(psCadena, "Cuentas CTS Movidas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    'Call CancelacionExternoCTS(psCadena, "Cuentas CTS Canceladas por Otras Agencias y/o CMAC's" & lsAdd, pnPagina, pnItem, lnTotDep, lnTotRet, psMoneda, psCodUser, psFecha)
'    'Call MovimientoTramiteCTS(psCadena, "Cuentas CTS Movidas en otras Agencias o CMAC's" & lsAdd, pnPagina, pnItem, psMoneda, psCodUser, psFecha)
'
'    psCadena = lsTEM + psCadena
'    lsTEM = ""
End Sub

Public Sub ReporteAperturasCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String)
    Dim Sql1 As String
    Dim RegTrandiaria As New ADODB.Recordset
    Dim lsDoc As String
    Dim lsCodCta As String
    Dim I As Long
    Dim lnTDep As Double
    Dim lnTRet As Double
    Dim lnValor As Double
    Dim lnPorRet As Double
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If psCodUser = "XXX" Then
        Sql1 = " Select M.cMovNro, MC.cOpeCod, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY CA.cCtaCod"
    Else
        Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
              & " nSaldoDisponible as nSaldDispAC, nTasaInteres nTasaIntAC" _
              & " From Mov M" _
              & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
              & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
              & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
              & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
              & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
              & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And Substring(M.cMovNro,18,2) = Substring(MC.cCtaCod,4,2) And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY CA.cCtaCod"
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
        
        psCadena = psCadena & Space(16 - Len(RegTrandiaria!cCodCta)) & RegTrandiaria!cCodCta _
                    & Space(20 - Len(Trim(lsDoc))) & Trim(lsDoc) _
                    & Space(22 - Len(Format(RegTrandiaria!nMonTran, "###,##0.00"))) & Format(RegTrandiaria!nMonTran, "###,##0.00") _
                    & Space(10 - Len(RegTrandiaria!cCodUsu)) & RegTrandiaria!cCodUsu _
                    & Space(17 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
                    & Space(22 - Len(Format(lnValor, "###,##0.00"))) & Format(lnValor, "###,##0.00") _
                    & Space(11 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
        
        pnItem = pnItem + 1
                
        If pnItem = lnRango Then
            psCadena = psCadena & oImpresora.gPrnSaltoPagina
            psCadena = psCadena & CabeceraPagina(psTitulo, pnPagina, pnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
            psCadena = psCadena & Encabezado("Nro Cuenta;16;Nro.Doc.;20;Monto;22;Usu.;10;Hora;17;Sald.Cnt;22;TInt;11;", pnItem)
        End If
        
        RegTrandiaria.MoveNext
    Wend
    
    If lnTRet <> 0 Or lnTDep <> 0 Then
        psCadena = psCadena & Encabezado("Resumen;16;" & Trim(Format(lnTDep, "###,##0.00")) & ";42; ;60;", pnItem)
        
        pnTotRet = pnTotRet + lnTRet
        pnTotDep = pnTotDep + lnTDep
        
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
    End If
End Sub

'Operacioens hechas por otras agencias en este servidor
Private Sub MovimientoExternosCTS(psCadena As String, psTitulo As String, pnPagina As Long, pnItem As Long, pnTotDep As Double, pnTotRet As Double, Optional psMoneda As String = "1", Optional psCodUser As String = "XXX", Optional psFecha As String = "", Optional psTipoOpe As String = "", Optional pbEsTramite As Boolean = False)
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
    Dim I As Long
    
    lnTRet = 0
    lnTDep = 0
    
    If pbEsTramite Then
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                 & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(MC.cCtaCod,4,2) cCodAge" _
                 & " From Mov M" _
                 & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                 & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                 & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                 & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                 & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                 & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                 & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                 & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                 & " And Substring(M.cMovNro,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(MC.cCtaCod,4,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(MC.cCtaCod,4,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,15,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(MC.cCtaCod,4,2), CA.cCtaCod"
        End If
    Else
        If psCodUser = "XXX" Then
            Sql1 = " Select M.cMovNro, MC.cOpeCod cCodOpe, MC.cCtaCod cCodCta, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion , Substring(M.cMovNro,18,2) cCodAge" _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
        Else
            Sql1 = " Select  M.cMovNro, MC.cOpeCod, MC.cCtaCod, cDocNro cNumDoc, abs(nMonto) nMontran," _
                  & " nSaldoDisponible as nSaldDispAC, nSaldoContable nSaldCnt, nTasaInteres nTasaIntAC, AGE.cAgeDescripcion, Substring(M.cMovNro,18,2) cCodAge " _
                  & " From Mov M" _
                  & " Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
                  & " Inner Join CaptacCTS CA On MC.cCtaCod = CA.cCtaCod" _
                  & " Inner Join Producto P On P.cCtaCod = CA.cCtaCod" _
                  & " Left  Join MovDoc MD On M.nMovNro = MD.nMovNro" _
                  & " Inner Join Agencias AGE On AGE.cAgeCod = Substring(M.cMovNro,18,2)" _
                  & " Where Substring(MC.cCtaCod,9,1) = '" & psMoneda & "'" _
                  & " And MC.cOpeCod in ('" & psTipoOpe & "')" _
                  & " And M.nMovFlag = " & MovFlag.gMovFlagVigente & "  And Substring(M.cMovNro,18,2) <> Substring(MC.cCtaCod,4,2)" _
                  & " And Substring(MC.cCtaCod,1,5) = '" & lsAgeCod & "' And M.cMovNro Like '" & psFecha & "%" & psCodUser & "' ORDER BY Substring(M.cMovNro,18,2), CA.cCtaCod"
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
            If lsGru <> "" And lsGru <> RegTrandiaria!cCodAge Then
                psCadena = psCadena & Encabezado("Total Agencia;15;" & Format(lnTDepAge, "###,##0.00") & ";57;" & Format(lnTRetAge, "###,##0.00") & ";20; ;46;", pnItem)
                lnTRetAge = 0
                lnTDepAge = 0
                lbBan = True
            End If

            If Len(RegTrandiaria!cCodCta) = 12 Then
                If lsGru <> RegTrandiaria!cCodAge Then
                    lsGru = RegTrandiaria!cCodAge
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
                                & Space(12 - Len(Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM"))) & Format(RegTrandiaria!dFectran, "hh:mm:ss AMPM") _
                                & Space(8 - Len(Format(RegTrandiaria!nTasaIntCTS, "###,##0.00"))) & Format(RegTrandiaria!nTasaIntCTS, "###,##0.00") & oImpresora.gPrnSaltoLinea
                               
            pnItem = pnItem + 1
            
            If pnItem = lnRango Then
                psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
        psCadena = psCadena & oImpresora.gPrnSaltoPagina
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
    Dim m As Integer, P As Integer
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
    
    lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0: lnRecs05 = 0: lnRecs06 = 0
    lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0: lnSdos05 = 0: lnSdos06 = 0
    
    lnCarLin = 170
    
    oCon.AbreConexion
    
    Set rsSdo = New ADODB.Recordset
    rsSdo.CursorLocation = adUseClient
    
    If psCodAge = "" Then
        sqlAgencia = ""
    Else
        sqlAgencia = " And Left(PR.cCtaCod,5) = '" & psCodAge & "'"
    End If
    
    lsRTFRfm = ""
    
    For m = 1 To 2
        lnCntPag = 1
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        lsMoneda = IIf(m = 1, "SOLES", "DOLARES")
        lsTitRp1 = "I N F O R M E   D E   S A L D O S   P O R   T I P O   D E   C U E N T A"
        lsTitRp2 = ""
        lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(38, " ") & "        A  H  O  R  R  O  S               P L A Z O    F I J O                     C    T    S                       T  O  T  A  L  " & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(38, " ") & "A C T I V A S        I N A C T I V A S" & String(29, " ") & "C L I E N T E S        E M P L E A D O S      A C U M U L A D O" & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(8, " ") & "C O N C E P T O S" & String(8, " ")
        lsRTFRfm = lsRTFRfm & "NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S NUMERO     S A L D O S" & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        For P = 0 To 3
            If (P < 4 Or P > 5) Then
                Select Case P
                    Case 0
                        sPers = "PERSONA NATURAL"
                    Case 1
                        sPers = "PERSONA JUR. SIN FINES DE LUCRO"
                    Case 2
                        sPers = "PERSONA JUR. CON FINES DE LUCRO"
                    Case 3
                        sPers = "CAJAS MUNICIPALES"
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
                    
                If P <> 3 Then
                              
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
                lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoTot)), 15, True, 9, 2) & oImpresora.gPrnSaltoLinea
                lnRecs06 = lnRecs06 + lnRecPro
                lnSdos06 = lnSdos06 + lnSdoTot
                  
                lnSdoTot = 0: lnRecPro = 0
            End If
           
        Next P
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
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
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos06)), 15, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
        If m < 2 Then
            lsRTFRfm = lsRTFRfm & oImpresora.gPrnSaltoPagina
            lsRTFRfm = lsRTFRfm & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0: lnRecs05 = 0: lnRecs06 = 0
            lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0: lnSdos05 = 0: lnSdos06 = 0
        End If
    Next m
    SdosTipCta = lsRTFRfm
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
    
    CabeRepo = CabeRepo & psCabe01 & oImpresora.gPrnSaltoLinea
    CabeRepo = CabeRepo & psCabe02 & oImpresora.gPrnSaltoLinea
    CabeRepo = CabeRepo & lsTitRp1 & oImpresora.gPrnSaltoLinea
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
                      
    Dim formato As String, I As Integer, lnPosDig As Integer
    
    If pnCampos = "0.00" Then
       JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
       Exit Function
    End If
    
    If pbComass Then
       lnPosDig = 0
       For I = 1 To pnDigEnt
           lnPosDig = lnPosDig + 1
           Select Case lnPosDig
              Case 1
                   formato = "0" & formato
              Case 4, 7, 10
                   formato = "#," & formato
              Case Else
                   formato = "#" & formato
           End Select
       Next I
       If pnDigDec > 0 Then
          formato = formato & "." & String(pnDigDec, "0")
       End If
    Else
       For I = 1 To pnDigEnt
           formato = IIf(I = 1, "0", "#") & formato
       Next I
       If pnDigDec > 0 Then
          formato = formato & "." & String(pnDigDec, "0")
       End If
    End If
    pnCampos = Format(pnCampos, formato)
    
    JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
    
End Function


Public Function EstratDep(Optional psCodAge As String = "") As String
    Dim m As Integer, P As Integer
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
    Dim I As Integer, j As Integer
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
        sqlAgencia = " And Left(PR.cCtaCod,5) = '" & psCodAge & "'"
    End If
    
    lnRecs01 = 0: lnRecs02 = 0: lnRecs03 = 0: lnRecs04 = 0
    lnSdos01 = 0: lnSdos02 = 0: lnSdos03 = 0: lnSdos04 = 0
    
    lnCarLin = 118
    lsLimite = "    100    500   1000   2000   5000  10000  25000  50000 1000009999999"
            
    oCon.AbreConexion
            
    For m = 1 To 2
       
        lsTipMon = Trim(Str(m))
        lnTCFHoy = IIf(lsTipMon = "1", 3.48, 1)
        lnCntPag = 1
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        lsMoneda = IIf(m = 1, "SOLES", "DOLARES")
        lsTitRp1 = "E S T R A T I F I C A C I O N   D E   L O S   D E P O S I T O S"
        lsTitRp2 = ""
        lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(32, " ") & " A  H  O  R  R  O  S   P L A Z O  F I J O       C    T    S          T  O  T  A  L" & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(7, " ") & "R  A  N  G  O  S" & String(9, " ")
        lsRTFRfm = lsRTFRfm & "NUMERO   S A L D O S  NUMERO   S A L D O S  NUMERO   S A L D O S  NUMERO   S A L D O S" & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        
        For P = 1 To 10
            lnLimInd = IIf(P > 1, P - 1, P)
            lnMinVal = IIf(P > 1, CCur(Mid$(lsLimite, (lnLimInd * 7) - 6, 7)), 0) * lnTCFHoy + 0.01
            lnMaxVal = CCur(Mid$(lsLimite, (P * 7) - 6, 7)) * lnTCFHoy
            lsLimInf = IIf(P > 1, "DE " & JDNum(Trim(Str(lnMinVal)), 12, True, 9, 2) & " A ", "HASTA.............")
            If P > 9 Then
               lnMinVal = CCur(Mid$(lsLimite, (lnLimInd * 7) - 6, 7)) * lnTCFHoy
               lsLimInf = "MAYOR A..........." & JDNum(Trim(Str(lnMinVal)), 12, True, 9, 2)
               lsLimSup = ""
            Else
               lsLimSup = JDNum(Trim(Str(lnMaxVal)), 12, True, 9, 2)
            End If
            
            lsRTFRfm = lsRTFRfm & lsLimInf & lsLimSup & "  "
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR" _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '232'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 12, True, 9, 2) & "  "
            lnRecs01 = lnRecs01 + rsQrySdo!nRecs:  lnSdos01 = lnSdos01 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:  lnSdoTot = lnSdoTot + lnSaldos
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR " _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '233'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 12, True, 9, 2) & "  "
            lnRecs02 = lnRecs02 + rsQrySdo!nRecs:   lnSdos02 = lnSdos02 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:   lnSdoTot = lnSdoTot + lnSaldos
            
            lsQrySdo = " SELECT COUNT(*) AS nRecs, IsNull(SUM(nSaldo),0) AS nSdos " _
                     & " FROM       Producto PR" _
                     & " WHERE      Left(nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "', '" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND SUBSTRING(cCtaCod,9,1) = '" & lsTipMon & "' And Substring(cCtaCod,6,3) = '234'" _
                     & " AND        nSaldo BETWEEN " & lnMinVal & " AND " & lnMaxVal & " " & sqlAgencia
                     
            If rsQrySdo.State = adStateOpen Then rsQrySdo.Close
            
            Set rsQrySdo = oCon.CargaRecordSet(lsQrySdo)
              
            lnSaldos = 0: lnSaldos = IIf(IsNull(rsQrySdo!nSdos) = True, 0, rsQrySdo!nSdos)
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(rsQrySdo!nRecs)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSaldos)), 12, True, 9, 2) & "  "
            lnRecs03 = lnRecs03 + rsQrySdo!nRecs:    lnSdos03 = lnSdos03 + lnSaldos
            lnRecPro = lnRecPro + rsQrySdo!nRecs:    lnSdoTot = lnSdoTot + lnSaldos
            
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecPro)), 6, False, 6, 0) & "  "
            lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoTot)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
            lnRecs04 = lnRecs04 + lnRecPro:    lnSdos04 = lnSdos04 + lnSdoTot
                  
            lnSdoTot = 0: lnRecPro = 0
        Next P
        
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & "T  O  T  A  L  E  S" & String(13, " ")
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs01)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos01)), 12, True, 9, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs02)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos02)), 12, True, 9, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs03)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos03)), 12, True, 9, 2) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnRecs04)), 6, False, 6, 0) & "  "
        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdos04)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina
           
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
                          psMoneda As String, psNumPag As String) As String

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
    
    CabeRepo1 = CabeRepo1 & psCabe01 & oImpresora.gPrnSaltoLinea
    CabeRepo1 = CabeRepo1 & psCabe02 & oImpresora.gPrnSaltoLinea
    CabeRepo1 = CabeRepo1 & lsTitRp1 & oImpresora.gPrnSaltoLinea
    CabeRepo1 = CabeRepo1 & lsTitRp2
        
End Function

Public Function LlenaEDAC(pnMoneda As Integer, pdFchPr1 As String, Optional pdFchPr2 As String) As String
    Dim rsQryEAC As ADODB.Recordset
    Set rsQryEAC = New ADODB.Recordset
    Dim lsQryEAC As String
    Dim ldFchIni As Date, ldFchFin As Date
    Dim lnCntMov As Integer
    Dim ldFchPro As Date
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    lsRTFDat = ""
    
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
           lsQryEAC = " SELECT cUltimoMovimiento,nMoneda,nAperturasNum,nAperturasMonto,nCancelacionesNum,nCancelacionesMonto,nRetiroIntacNum,nRetiroIntacMon,nCancelacionesInacNum,nCancelacionesInacMonto,nDepositosNum,nDepositosMonto,nRetiroNum,nRetiroMonto,nSaldoAnterior,nSaldo,nSaldoCMAC,nChequeCMAC,nCuentasVigentes,nChequesMonto,nInteresCapitalizados,nChequeValorizacionMonto,nChequevalorizacionNum,nNumCMAC FROM EstadisticaAhorro" _
                    & " WHERE nMoneda = '" & Trim(Str(pnMoneda)) & "' AND cUltimoMovimiento BETWEEN '" & Format(CDate(ldFchPro), "yyyymmdd") & "' AND '" & Format(CDate(ldFchPro) + 1, "yyyymmdd") & "' AND Substring(cUltimoMovimiento,15,5) = '" & lsAgeCod & "' Order by cUltimoMovimiento "
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If rsQryEAC.EOF And rsQryEAC.BOF Then
              rsQryEAC.Close
              LlenaEDAC = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For I = 2 To 14
                  laMovDia(I, lnCntMov) = 0
              Next I
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
                laMovDia(14, lnCntMov) = (!nSaldo + !nChequeValorizacionMonto + !nChequeCMAC)
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop
        PrvEDAC pnMoneda
        
        LlenaEDAC = lsRTFDat
End Function

Public Function PrvEDAC(pnMoneda As Integer)
    Dim lnCarLin As Integer, I As Integer
    Dim lnCabece As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCntArr As Integer
    Dim lsMoneda As String
    
    lnCabece = 0
    lnCarLin = 136
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.Page)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        lsTitRep = "Movimiento Diario de Ahorro Corriente"
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & lsCabe02 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & "    M O V I M I E N T O S" & Space(27) & "M   O   N   T   O   S" & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & "   SALDO A.C." & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea            ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For I = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, I) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, I), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(12, I), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(13, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(14, I), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea   '  Format(laMovDia(14, I), "@@@@@@@@@@@@") & oImpresora.gPrnSaltoLinea
        lnTotA05 = laMovDia(5, I)
    Next I
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
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
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA13), 11, True, 8, 2) & oImpresora.gPrnSaltoLinea   '   " "
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & JDNum((lnTotA14 / lnCntArr), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina                ' hasta 160 caracteres x línea
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
           lsQryEAC = " SELECT cUltimoMovimiento,nMoneda,nAperturasNum,nAperturasMonto,nCancelacionesNum,nCancelacionesMonto,nMovimientosNum,nMovimientosMonto,nCuentasVigentes,nIntnteresCapitalizado,nSaldoAnterior,nSaldo,nChequeValorizacionMonto,nChequevalorizacionNum,nNumCMAC,nSaldoCMAC FROM Estadisticaplazofijo  " _
                    & " WHERE nMoneda = '" & Trim(Str(pnMoneda)) & "' AND cUltimoMovimiento BETWEEN '" & Format(CDate(ldFchPro), "yyyymmdd") & "' AND '" & Format(CDate(ldFchPro) + 1, "yyyymmdd") & "' AND Substring(cUltimoMovimiento,15,5) = '" & lsAgeCod & "' Order by cUltimoMovimiento "
           
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If RSVacio(rsQryEAC) Then
              rsQryEAC.Close
              Set rsQryEAC = Nothing
              LlenaEDPF = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For I = 2 To 11
                  laMovDia(I, lnCntMov) = 0
              Next I
           Else
           
           With rsQryEAC
                           
                laMovDia(1, lnCntMov) = Format(!dEstadPF, gsFormatoFechaView)
                laMovDia(2, lnCntMov) = !nNumAperPF
                laMovDia(3, lnCntMov) = !nNumCancPF
                laMovDia(4, lnCntMov) = !nNumMovPF
                laMovDia(5, lnCntMov) = !nNumVigPF
                laMovDia(6, lnCntMov) = !nMonAperPF
                laMovDia(7, lnCntMov) = !nMonCancPF
                laMovDia(8, lnCntMov) = !nMonMovPF
                laMovDia(9, lnCntMov) = !nIntCapPF
                laMovDia(10, lnCntMov) = (!nSaldoPF + !nMonChqVal)
                laMovDia(11, lnCntMov) = 0
                
                lnTotA02 = lnTotA02 + !nNumAperPF
                lnTotA03 = lnTotA03 + !nNumCancPF
                lnTotA04 = lnTotA04 + !nNumMovPF
                lnTotA05 = lnTotA05 + !nNumVigPF
                lnTotA06 = lnTotA06 + !nMonAperPF
                lnTotA07 = lnTotA07 + !nMonCancPF
                lnTotA08 = lnTotA08 + !nMonMovPF
                lnTotA09 = lnTotA09 + !nIntCapPF
                lnTotA10 = lnTotA10 + (!nSaldoPF + !nMonChqVal)
                lnTotA11 = 0
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop
        
        'If CierreRealizado(2) Then
           lbCieMes = True
        'Else
        '   lbCieMes = False
        'End If
        
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
    Dim lnCarLin As Integer, I As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCabece As Integer
    Dim lnCntArr As Integer
    
    lnCabece = 0
    lnCarLin = 110
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.Page)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        lsTitRep = "Movimiento Diario de Plazo Fijo"
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & lsCabe02 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & "       M O V I M I E N T O S" & Space(21) & "M   O   N   T   O   S" & String(16, " ")
    lsRTFDat = lsRTFDat & "  SALDOS PF " & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & "  INT.DEVENG." & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea            ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For I = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, I) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, I), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, I), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
        lnTotA05 = laMovDia(5, I)
    Next I
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "  TOTAL    "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA02), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA03), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA04), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Int(lnTotA05), 5, False, 5, 0) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA06), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA07), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA08), 11, True, 8, 2) & " "
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA09), 11, True, 8, 2) & oImpresora.gPrnSaltoLinea
'    lsRTFDat = lsRTFDat & JDNum(lnTotA10, 12, True) & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & JDNum((lnTotA11 / lnCntArr), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina                ' hasta 160 caracteres x línea
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
    
    ldFchIni = CDate(pdFchPr1)
    ldFchFin = CDate(pdFchPr2)
    
    oCon.AbreConexion
    
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
                    & " WHERE nMoneda = '" & Trim(Str(pnMoneda)) & "' AND cUltimoMovimiento BETWEEN '" & Format(CDate(ldFchPro), "yyyymmdd") & "' AND '" & Format(CDate(ldFchPro) + 1, "yyyymmdd") & "' AND Substring(cUltimoMovimiento,15,5) = '" & lsAgeCod & "' Order by cUltimoMovimiento "
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           Set rsQryEAC = oCon.CargaRecordSet(lsQryEAC)
       
           If RSVacio(rsQryEAC) Then
              rsQryEAC.Close
              Set rsQryEAC = Nothing
              LlenaEDCTS = ""
              laMovDia(1, lnCntMov) = Format(ldFchPro, gsFormatoFechaView)
              For I = 2 To 14
                  laMovDia(I, lnCntMov) = 0
              Next I
'              Exit Function
           Else
           
           With rsQryEAC
                           
                laMovDia(1, lnCntMov) = Format(!dEstadCTS, gsFormatoFechaView)
                laMovDia(2, lnCntMov) = !nNumAperCTS
                laMovDia(3, lnCntMov) = !nNumCancCTS
                laMovDia(4, lnCntMov) = !nNumVigCTS
                laMovDia(5, lnCntMov) = !nNumDepCTS
                laMovDia(6, lnCntMov) = !nNumRetCTS
                laMovDia(7, lnCntMov) = !nMonAperCTS
                laMovDia(8, lnCntMov) = !nMonCancCTS
                laMovDia(9, lnCntMov) = !nMonDepCTS
                laMovDia(10, lnCntMov) = !nMonRetCTS
                laMovDia(11, lnCntMov) = !nIntCapCTS
                laMovDia(12, lnCntMov) = (!nSaldoCTS + !nMonChqVal)
                laMovDia(13, lnCntMov) = !nSaldEmp
                laMovDia(14, lnCntMov) = 0
                
                lnTotA02 = lnTotA02 + !nNumAperCTS
                lnTotA03 = lnTotA03 + !nNumCancCTS
                lnTotA04 = lnTotA04 + !nNumVigCTS
                lnTotA05 = lnTotA05 + !nNumDepCTS
                lnTotA06 = lnTotA06 + !nNumRetCTS
                lnTotA07 = lnTotA07 + !nMonAperCTS
                lnTotA08 = lnTotA08 + !nMonCancCTS
                lnTotA09 = lnTotA09 + !nMonDepCTS
                lnTotA10 = lnTotA10 + !nMonRetCTS
                lnTotA11 = lnTotA11 + !nIntCapCTS
                lnTotA12 = lnTotA12 + (!nSaldoCTS + !nMonChqVal)
                lnTotA13 = lnTotA13 + !nSaldEmp
                lnTotA14 = 0
                
           End With
           End If
           ldFchIni = DateAdd("d", 1, ldFchPro)
        Loop

        'If CierreRealizado(2) Then
        '   lbCieMes = True
        'Else
           lbCieMes = False
        'End If
        
        If lbCieMes Then
            ldFchIni = CDate(pdFchPr1)
            ldFchFin = CDate(pdFchPr2)
            'lsQryEAC = " SELECT SUM(ABS(nMonTran)) AS nMonPro FROM TransAho " _
                     & " WHERE  cCodOpe = '" & gsCTSProInt & "' AND (cFlag IS NULL OR cFlag <> 'X') " _
                     & " and    substring(cCodCta,6,1) = '" & Trim(Str(pnMoneda)) & "' " _
                     & " AND    (dFecTran BETWEEN '" & Format(ldFchIni, gsFormatoFecha) & "' AND '" & Format(DateAdd("d", 1, ldFchFin), gsFormatoFecha) & "') "
           If rsQryEAC.State = adStateOpen Then rsQryEAC.Close
           
           'rsQryEAC.Open lsQryEAC, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
        
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
    
    Dim lnCarLin As Integer, I As Integer
    Dim lsTitRep As String, lsCabe01 As String, lsCabe02 As String
    Dim lnCabece As Integer
    Dim lnCntArr As Integer
    
    lnCabece = 0
    lnCarLin = 142
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    
    '   Definición de Cabecera 1
        lsCabe01 = UCase(Trim(lsNomAge)) & " - " & lsMoneda
        lsCabe01 = lsCabe01 & Space((lnCarLin - 34) - Len(lsCabe01))
        lsCabe01 = lsCabe01 & "PAGINA: " & FillNum(Trim(Str(Printer.Page)), 4, " ")
        lsCabe01 = lsCabe01 & Space(5) & "FECHA: " & Format(ldFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
        lsCabe02 = "SECCION AHORRO"
        lsCabe02 = lsCabe02 & Space((lnCarLin - 17) - Len(lsCabe02))
        lsCabe02 = lsCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
        lsTitRep = "Movimiento Diario de C T S "
        lsTitRep = String(Int((lnCarLin - Len(lsTitRep)) / 2), " ") & lsTitRep
    
    lsRTFDat = lsCabe01 & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & lsCabe02 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & UCase(lsTitRep) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & "    M O V I M I E N T O S" & Space(27) & "M   O   N   T   O   S" & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & "   PROVISIONES" & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea            ' hasta 160 caracteres x línea
            
'       Línea de Impresión
    lnCntArr = UBound(laMovDia, 2)
    For I = 1 To lnCntArr
        lsRTFDat = lsRTFDat & laMovDia(1, I) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(2, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(3, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(4, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(5, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(6, I), 5, False, 5, 0) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(7, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(8, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(9, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(10, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(11, I), 11, True, 8, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(12, I), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(13, I), 13, True, 10, 2) & " "
        lsRTFDat = lsRTFDat & JDNum(laMovDia(14, I), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea   '  Format(laMovDia(14, I), "@@@@@@@@@@@@") & oImpresora.gPrnSaltoLinea
        lnTotA05 = laMovDia(5, I)
    Next I
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
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
    lsRTFDat = lsRTFDat & JDNum(Str(lnTotA11), 11, True, 8, 2) & oImpresora.gPrnSaltoLinea
'    lsRTFDat = lsRTFDat & Format(Format(lnTotA13, "###,##0.00"), "@@@@@@@@@@") & oImpresora.gPrnSaltoLinea   '   " "
'    lsRTFDat = lsRTFDat & Format(Format(lnTotA14, "#,###,##0.00"), "@@@@@@@@@@@@") & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
    lsRTFDat = lsRTFDat & "PROMEDIO " & oImpresora.gPrnSaltoLinea
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
    lsRTFDat = lsRTFDat & JDNum((lnTotA14 / lnCntArr), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsRTFDat = lsRTFDat & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina               ' hasta 160 caracteres x línea
End Function

Public Function GenSdosFM(psTipAho As String, psTipPer As String, psTipMon As String, _
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
    
    oCon.AbreConexion
    
    lsRTFRfm = ""
    
    lsMoneda = IIf(psTipMon = "1", "SOLES", "DOLARES")
    
    Select Case psTipPer
        Case "0"
             lsPerson = "PERSONA(S)  NATURAL(ES)" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case "1"
             lsPerson = "PERSONA(S)  JURIDICA(S)  SIN  FINES  DE  LUCRO" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case "2"
             lsPerson = "PERSONA(S)  JURIDICA(S)  CON  FINES  DE  LUCRO" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
        Case "3"
             lsPerson = "PERSONA(S)  CMAC" & IIf(psTipAho = Producto.gCapAhorros, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
    End Select
        
    lsQrySdo = ""
    Select Case psTipAho
        Case Producto.gCapAhorros
   
             lsTipAho = "A H O R R O S"
             lsQrySdo = " SELECT PR.cCtaCod cCodCta, bOrdPag cOrdPag, CA.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt, CA.nIntAcum As nIntMes, CAA.nSaldoAnterior AS nSdoAnt, CAA.dUltContacto AS dUltMovAC, PE.cPersnombre cNomPers, PP.cPersCod cCodPers FROM Producto PR" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod" _
                      & " INNER JOIN Persona PE ON Pp.cPersCod = PE.cPersCod" _
                      & " INNER JOIN Captaciones CA ON PR.cCtaCod = CA.cCtaCod" _
                      & " INNER JOIN CaptacAhorros CAA ON PR.cCtaCod = CAA.cCtaCod" _
                      & "" _
                      & " WHERE Left(PR.nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND PE.nPersPersoneria = '" & psTipPer & "'" _
                      & " AND PP.nPrdPersRelac = 10 AND CAA.bInactiva = " & pbCtaIna & "" _
                      & " AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' " _
                      & " ORDER BY AC.cOrdPag, AC.cCodCta"
        Case Producto.gCapPlazoFijo
             lsTipAho = "P L A Z O   F I J O"
             lsQrySdo = " SELECT CA.dApertura dAperPF, DateAdd(day, CAPF.nPlazo, CAPF.dRenovacion) AS dFchVct, CAPF.nPlazo, PR.cCtaCod cCodCta, PE.cPersNombre cNomPers, CA.nSaldoDisp AS nCapital, " _
                      & " CAPF.nIntPag, CA.nIntAcum, CAPF.nApertura, PR.nTasaInteres nTasaIntPF, CAPF.dAuxiliar AS dUltMovPF, PE.cPersCod " _
                      & " FROM Producto PR" _
                      & " Inner Join CaptacPlazoFijo CAPF On PR.cCtaCod = CAPF.cCtaCod INNER JOIN Captaciones CA ON CAPF.cCtaCod = CA.cCtaCod " _
                      & " INNER JOIN ProductoPersona PP ON CAPF.cCtaCod = PP.cCtaCod INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " _
                      & " WHERE Left(PR.nPrdEstado,2) NOT IN ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') AND PE.nPersPersoneria = '" & psTipPer & "' " _
                      & " AND PP.nPrdPersRelac = 10 AND SUBSTRING(CAPF.cCtaCod,9,1) =  '" & psTipMon & "'" _
                      & " AND SUBSTRING(CAPF.cCtaCod,9,1) = '" & psTipMon & "'  " _
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
                      & " CA.dUltCierre AS dUltMovCTS, 0 As nUltDep, PR.nSaldo As nCapital, CA.nIntAcum AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers " _
                      & " FROM Producto PR  " _
                      & " INNER JOIN Captaciones CA ON CA.cCtaCod = PR.cCtaCod  INNER JOIN CaptacCTS CACTS ON CACTS.cCtaCod = PR.cCtaCod" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod  INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod" _
                      & " WHERE Left(PR.nPRDEstado,2) NOT IN ('13','14') AND LTRIM(RTRIM(CACTS.cCodInst)) = '" & lsInstit & "'" _
                      & " AND PP.nPrdPersRelac = 10 AND SUBSTRING(CACTS.cCtaCod,9,1) = '" & psTipPer & "'" _
                      & " AND SUBSTRING(CACTS.cCtaCod,9,1) = '" & psTipMon & "'  " _
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


Private Sub PrtSdosFM(ByVal prQryDat As ADODB.Recordset, psMoneda As String, psTipAho As String, psPerson As String)

    Dim lnCarLin As Integer, I As Integer
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
    
    'If CierreRealizado(2) Then
       lbCieMes = True
    'Else
    '   lbCieMes = False
    'End If
    
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
                    lsTitRp2 = IIf(!cOrdPag = "N", "", " CON ORDEN DE PAGO  -  ") & psPerson
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
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                     lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
                  Else
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
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
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
                        lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
                     Else
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
                        lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
                     End If
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, psMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
               
               If Mid$(!cCodCta, 6, 3) = Producto.gCapAhorros Then
                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "  CUENTA  " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "     INTERES  " & " "
                  lsRTFRfm = lsRTFRfm & "  SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
                  lsRTFRfm = lsRTFRfm & " ULTIMO  " & oImpresora.gPrnSaltoLinea
                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "      DEL CLIENTE     " & String(10, " ")
                  lsRTFRfm = lsRTFRfm & "    DEL MES  " & " "
                  lsRTFRfm = lsRTFRfm & "DISPONIBLE " & " "
                  lsRTFRfm = lsRTFRfm & "   CONTABLE  " & " "
                  lsRTFRfm = lsRTFRfm & "MOVIMIENTO" & oImpresora.gPrnSaltoLinea
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
                  lsRTFRfm = lsRTFRfm & " ULTIMO" & oImpresora.gPrnSaltoLinea
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
                  lsRTFRfm = lsRTFRfm & "CONTACTO" & oImpresora.gPrnSaltoLinea
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
                  lsRTFRfm = lsRTFRfm & "    ULTIMO" & oImpresora.gPrnSaltoLinea
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
                  lsRTFRfm = lsRTFRfm & "  CONTACTO" & oImpresora.gPrnSaltoLinea
               End If
               
               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea           ' hasta 160 caracteres x línea
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
                    lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & oImpresora.gPrnSaltoLinea
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
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovAC, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
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
                       lnCalInt = GetInteres(!nTasaIntPF, !nIntAcum + (!nIntAcum + !nApertura), lnNumDia, False)
                       
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
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovPF, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
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
                       lsRTFRfm = lsRTFRfm & Format(!dUltMovCTS, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
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
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
          lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
          lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
       Else
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
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
               lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cPersNombre), False), 34) & oImpresora.gPrnSaltoLinea
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


Public Function GeneraSaldos(psTipAho As String, psTipPer As String, psTipMon As String, _
                             psFchPro As String) As String
    Dim lsQrySdo As String
    Dim rsQrySdo As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    lsRTFSdo = ""
    lsMoneda = IIf(psTipMon = "1", "SOLES", "DOLARES")
    Select Case psTipPer
        Case "0"
             lsPerson = "PERSONA NATURAL"
        Case "1"
             lsPerson = "PERSONA JURIDICA SIN FINES DE LUCRO"
        Case "3"
             lsPerson = "PERSONA JURIDICA CON FINES DE LUCRO"
    End Select
    
    lsQrySdo = ""
    Select Case psTipAho
        Case Producto.gCapAhorros
             lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAPAHO.bOrdPag cOrdPag, CAP.nSaldoDisp AS nSdoDis, PR.nSaldo AS nSdoCnt, " _
                      & " CAP.nIntAcum AS nIntMes, PE.cPersNombre cNomPers, PE.cPersCod cCodPers, CAPAHO.dUltContacto AS dUltMovAC," _
                      & " PR.nTasaInteres nTasaIntAC, CAPAHO.nSaldoAnterior AS nSdoAnt" _
                      & " FROM Producto PR" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod" _
                      & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod" _
                      & " INNER JOIN CaptacAhorros CAPAHO ON PR.cCtaCod = CAPAHO.cCtaCod" _
                      & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod" _
                      & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                      & " AND CAP.nPersoneria = '" & psTipPer & "' AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' And Left(PR.cCtaCod ,5) = '" & lsAgeCod & "'" _
                      & " ORDER BY CAPAHO.bOrdPag, PR.cCodCta"
        Case Producto.gCapPlazoFijo
             lsQrySdo = " SELECT PR.cCtaCod cCodCta, CAP.nSaldoDisp AS nCapital, CAP.nIntAcum AS nIntDev, " _
                      & " PR.nSaldo AS nSdoCnt, PE.cPersNombre cNomPers, PP.cPersCod cCodPers, PR.nTasaInteres," _
                      & " CAPPF.dAuxiliar AS dUltMovPF" _
                      & " FROM Producto PR" _
                      & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod" _
                      & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod" _
                      & " INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod" _
                      & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod" _
                      & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                      & " AND CAP.nPersoneria = '" & psTipPer & "' AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' And Left(PR.cCtaCod ,5) = '" & lsAgeCod & "'" _
                      & " ORDER BY PR.cCodCta"
        Case Producto.gCapCTS
             If psTipPer = "0" Then
                lsQrySdo = " SELECT CAPCTS.cCtaCod cCodCta, CAPCTS.nSaldRetiro As nSdoDis, PR.nSaldo nSdoCnt, " _
                         & " CAP.nSaldoDisp As nCapital, CAP.dUltCierre AS dUltMovCTS, CAPCTS.cCodInst," _
                         & " PE.cPersNombre cNomPers, PP.cPersCod cCodPers" _
                         & " FROM Producto PR" _
                         & " INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod" _
                         & " INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod" _
                         & " INNER JOIN CaptacCTS CAPCTS ON PR.cCtaCod = CAPCTS.cCtaCod" _
                         & " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod" _
                         & " WHERE Left(PR.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') " _
                         & " AND CAP.nPersoneria = '" & psTipPer & "' AND  PP.nPrdPersRelac = '10' AND SUBSTRING(PR.cCtaCod,9,1) = '" & psTipMon & "' And Left(CAPCTS.cCtaCod ,5) = '" & lsAgeCod & "'" _
                         & " ORDER BY CAPCTS.cCodInst, PR.cCodCta"
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
    Else
       PrtSaldos rsQrySdo
       GeneraSaldos = lsRTFSdo
    End If
End Function

Private Function PrtSaldos(prQryDat As ADODB.Recordset)

    Dim lnCarLin As Integer, I As Integer
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnTotDep As Currency, lnTotRet As Currency
    Dim lsCodPro As String, lsNumPag As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lsSubTit As String, lnCntCta As Integer
    Dim lnSdoCnt As Currency, lnSdoDis As Currency
    Dim lnCapita As Currency, lsCodCta As String
    Dim lnCapiFi As Currency, lnSdDiFi As Currency, lnSdCnFi As Currency
    Dim lnCtaPag As Integer, lnTotCta As Integer
    Dim lnIntMes As Currency, lnSdInFi As Currency
    Dim lnCntRel As Integer
    Dim lbCtaIde As Boolean
    Dim lnIntCap As Currency
    Dim lnIntDev As Currency
    Dim lbCieMes As Boolean
    
    'If CierreRealizado(2) Then
    '    lbCieMes = True
    'Else
    '    lbCieMes = False
    'End If
    
    lsNumPag = ""
    lnLinPag = 65
    lnCarLin = 105
    lsTitRp1 = "REPORTE DE SALDO DE CUENTA(S)"
    lsTitRp2 = ""
    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0
    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0
    lnCntPag = 0
    lbCtaIde = False
    
    With prQryDat
         Do While Not .EOF
            Select Case Mid$(!cCodCta, 6, 3)
               Case Producto.gCapAhorros
                    lnCarLin = 116
                    lsTitRp2 = "AHORROS" & IIf(!cOrdPag = "N", "", " CON ORDEN DE PAGO") & " - " & lsPerson
               Case Producto.gCapPlazoFijo
                    lnCarLin = 103
                    lsTitRp2 = "PLAZO FIJO" & " - " & lsPerson
               Case Producto.gCapCTS
                    lnCarLin = 116
                    lsPerson = "PERSONA NATURAL - " & IIf(LTrim(RTrim(!cCodInst)) = gsInstCmac, "EMPLEADOS", "CLIENTES")
                    lsTitRp2 = "C T S" & " - " & lsPerson
            End Select
            
            If lsTitRp2 <> lsCodPro Then
               If lsNumPag <> "" Then
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
                  lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(17, " ")
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
                  lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
                  If Left(lsCodPro, 7) = "AHORROS" Then
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                  Else
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                  End If
                  lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
                  lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0
                  lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0
               End If
               lsCodPro = lsTitRp2
               lnLinPag = 65
            End If
            
            If lnLinPag > 58 Then
               If lsNumPag <> "" Then
                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
                     lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ")
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
                     lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
                     If Left(lsCodPro, 7) = "AHORROS" Then
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     Else
                        lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                     End If
                     lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0
                  End If
               End If
               
               lnCntPag = lnCntPag + 1
               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
               lsRTFSdo = lsRTFSdo & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
               If Mid$(!cCodCta, 6, 3) = Producto.gCapPlazoFijo Then
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   INTERES   "
                  lsRTFSdo = lsRTFSdo & "   SALDO    " & " "
                  lsRTFSdo = lsRTFSdo & "    ULTIMO" & oImpresora.gPrnSaltoLinea
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & "  DEVENGADO  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE  " & " "
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & oImpresora.gPrnSaltoLinea
               Else
                  lsRTFSdo = lsRTFSdo & "   NUMERO    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "NOMBRE / RAZON SOCIAL" & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "             "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   SALDO     "
                  lsRTFSdo = lsRTFSdo & "   INTERESES" & " "
                  lsRTFSdo = lsRTFSdo & "    ULTIMO" & oImpresora.gPrnSaltoLinea
                  lsRTFSdo = lsRTFSdo & "   CUENTA    " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "     DEL CLIENTE     " & String(10, " ")
                  lsRTFSdo = lsRTFSdo & "   CAPITAL   "
                  lsRTFSdo = lsRTFSdo & " DISPONIBLE  "
                  lsRTFSdo = lsRTFSdo & "  CONTABLE   "
                  lsRTFSdo = lsRTFSdo & "    DEL  MES" & " "
                  lsRTFSdo = lsRTFSdo & "MOVIMIENTO" & oImpresora.gPrnSaltoLinea
               End If
               lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea           ' hasta 160 caracteres x línea
               lnLinPag = 9
            End If
            lnCntRel = 0
            If !cCodCta = lsCodCta Then
               If lbCtaIde = False Then
                  lnCntRel = RelCta(!cCodCta, !cCodPers)
                  lbCtaIde = True
               End If
            Else
               lsCodCta = !cCodCta
               lbCtaIde = False
               lsRTFSdo = lsRTFSdo & !cCodCta & " "
               If Len(Trim(!cNomPers)) < 35 Then
                  lsRTFSdo = lsRTFSdo & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 35), 35, " ") & " "
               Else
                  lsRTFSdo = lsRTFSdo & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 35) & " "
               End If
               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
               
               Select Case Mid$(!cCodCta, 6, 3)
                  Case Producto.gCapAhorros
                       lsRTFSdo = lsRTFSdo & JDNum("0.00", 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       
                       lnIntCap = 0
                       If lbCieMes Then
                            lnIntCap = !nSdoDis - !nSdoAnt
                       Else
                            lnIntCap = !nIntMes
                       End If
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovAC, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
                       lnCapita = lnCapita + 0: lnCapiFi = lnCapiFi + 0
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdoDis = lnSdoDis + !nSdoDis
                       lnIntMes = lnIntMes + lnIntCap
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnSdDiFi = lnSdDiFi + !nSdoDis
                       lnSdInFi = lnSdInFi + lnIntCap
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapPlazoFijo
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 12, True, 9, 2) & " "
                       lnIntDev = !nIntDev
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovPF, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoDis = lnSdoDis + lnIntDev
                       lnSdDiFi = lnSdDiFi + lnIntDev
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
                       lnCntRel = RelCta(!cCodCta, !cCodPers)
                       lbCtaIde = True
                  Case Producto.gCapCTS
                       lsRTFSdo = lsRTFSdo & JDNum(!nCapital, 11, True, 8, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoDis, 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & JDNum(!nSdoCnt, 12, True, 9, 2) & " "
                       
                       lnIntCap = 0
'                       If lbCieMes Then
'                            lnIntCap = !nSdoDis - !nSdoAnt
'                       Else
'                            lnIntCap = !nIntMes
'                       End If
                       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
                       lsRTFSdo = lsRTFSdo & Format(!dUltMovCTS, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
                       
                       lnCapita = lnCapita + !nCapital
                       lnCapiFi = lnCapiFi + !nCapital
                       lnSdoCnt = lnSdoCnt + !nSdoCnt
                       lnSdCnFi = lnSdCnFi + !nSdoCnt
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
            
            .MoveNext
         Loop
    End With
    If prQryDat.EOF Then
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(14, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
       Else
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
       lsRTFSdo = lsRTFSdo & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(17, " ")
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
       lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
       If Left(lsCodPro, 7) = "AHORROS" Then
         lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & " "
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdInFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
       Else
          lsRTFSdo = lsRTFSdo & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
       End If
       lsRTFSdo = lsRTFSdo & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina              ' hasta 160 caracteres x línea
       lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0: lnIntMes = 0
       lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0: lnSdInFi = 0
    End If
End Function

Public Function RepEstratPlazoFijo() As String
    Dim lsReporte As String
    Dim rsPlazo As ADODB.Recordset
    Set rsPlazo = New ADODB.Recordset
    Dim nPlazos As Integer
    Dim j As Integer
    Dim nRangos(1 To 6, 1 To 2) As Integer
    Dim sRangos(1 To 6) As String
    Dim nNum As Long, nNumCTS As Long
    Dim nMonto As Currency, nMontoCTS As Currency
    Dim lsNum As String, lsMonto As String
    Dim nTotalNum As Long
    Dim nTotalMonto As Currency
    Dim lnCarLin As Integer, I As Integer
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
    lsReporte = CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, "", lsNumPag) & oImpresora.gPrnSaltoLinea
    lsReporte = lsReporte & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea               ' hasta 160 caracteres x línea
    
    For j = 1 To 2
        nTotalNum = 0
        nTotalMonto = 0
        If j = 1 Then
            lsReporte = lsReporte & "MONEDA NACIONAL" & oImpresora.gPrnSaltoLinea
        Else
            lsReporte = lsReporte & "MONEDA EXTRANJERA" & oImpresora.gPrnSaltoLinea
        End If
        lsReporte = lsReporte & String(17, "-") & oImpresora.gPrnSaltoLinea
        lsReporte = lsReporte & "PLAZO" & Space(20) & "NUMERO" & Space(18) & "MONTO" & oImpresora.gPrnSaltoLinea
        For I = 1 To nPlazos
            nNum = 0
            nMonto = 0
            VSQL = " Select Count(PR.cCtaCod) Num, Sum(nSaldo) Monto" _
                 & " From CaptacPlazoFijo PF " _
                 & " Inner Join Producto PR On PF.cCtaCod = PR.cCtaCod Where " _
                 & " nPlazo between " & nRangos(I, 1) & " and " & nRangos(I, 2) & " and " _
                 & " Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')" _
                 & " And Substring(PR.cCtaCod,9,1) = '" & j & "'  And Left(PR.cCtaCod,5) = '" & lsAgeCod & "'"
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
            
            If I = 6 Then
                GetSaldoCTS j, nNumCTS, nMontoCTS
                nNum = nNum + nNumCTS
                nMonto = nMonto + nMontoCTS
            End If
            nTotalNum = nTotalNum + nNum
            nTotalMonto = nTotalMonto + nMonto
            lsNum = Str(nNum)
            lsMonto = Format$(nMonto, "#,##0.00")
            lsReporte = lsReporte & sRangos(I) & Space(14) & JDNum(lsNum, 5, False, 5, 0) & Space(10) & JDNum(lsMonto, 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
        Next I
        lsReporte = lsReporte & "TOTAL       " & Space(14) & JDNum(Str(nTotalNum), 5, False, 5, 0) & Space(10) & JDNum(Str(nTotalMonto), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
        lsReporte = lsReporte & oImpresora.gPrnSaltoLinea
    Next j
    Set rsPlazo = Nothing
    RepEstratPlazoFijo = lsReporte
End Function

Private Sub GetSaldoCTS(ByVal psMoneda As String, ByRef pnNum As Long, ByRef pnMonto As Currency)
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
         & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "' And Left(PR.cCtaCod,5) = '" & lsAgeCod & "'"
    
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
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS SOLES - CON ORDEN PAGO -", "1", "232", 0, 0, 1, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS SOLES - SIN ORDEN PAGO -", "1", "232", 0, 0, 0, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS DOLARES - CON ORDEN PAGO -", "2", "232", 0, 0, 1, True)
    lsCad = lsCad & GetCtaInaAho("AHORRO INACTIVAS DOLARES - SIN ORDEN PAGO -", "2", "232", 0, 0, 0, True)
    RepCtaIna = lsCad
End Function

Public Function GetCtaInaAho(psTit As String, psMoneda As String, psProd As String, lnPagina As Long, lnItem As Long, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As ADODB.Recordset
    Set rsAho = New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 48
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
        lsCadena = lsCadena & Encabezado("ITEM;4;CUENTA;20;NOMBRE;25; ;25;SALDO;12; ;2;FEC.APE;10; ;2;ULT.CNT;10;", lnItem)
        
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
                lsUltCnt = Format$(rsAho!UltCon, gsFormatoFechaView)
            
            Else
                lsContador = ""
                lnMonto = 0
                lsFecha = ""
                lsUltCnt = ""
            End If
            lsCodCta = IIf(rsAho!Cuenta = lsCodCtaAnt, "", rsAho!Cuenta)
            lsNomPers = PstaNombre(rsAho!Nombre)
            lsMonto = IIf(lnMonto = 0, "", JDNum(Str(lnMonto), 12, True, 9, 2))
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(2) & lsMonto & Space(2) & lsFecha & Space(2) & lsUltCnt & oImpresora.gPrnSaltoLinea
            lnAcumItem = lnAcumItem + 1
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("ITEM;4;CUENTA;20;NOMBRE;25; ;25;SALDO;12; ;2;FEC.APE;10; ;2;ULT.CNT;10;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & String(106, "=") & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & lsContador & Space(20) & "TOTAL SALDO" & Space(35) & JDNum(Trim(Str(lnAcum)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & String(106, "=") & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If
    
    GetCtaInaAho = lsCadena
    
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
    lsCad = lsCad & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "CAJA MUNICIPAL DE " & Space(50) & ArmaFecha(ldFecSis) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "    TRUJILLO      " & Space(50) & Str(Time()) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Space(30) & "RESUMEN DE DESCUENTO DE INACTIVAS" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Space(30) & "---------------------------------" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Agencia : " & lsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For I = 1 To 2
        lnNum1 = 0: lnMon1 = 0: lnNum2 = 0: lnMon2 = 0: lnNum3 = 0
        lnMon3 = 0: lnTotNumero = 0: lnTotMonto = 0
        Sql1 = " SELECT COUNT(MC.nMovNro) as Numero ,ISNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoEstActInac & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & I & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
             & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
        
        Sql2 = " SELECT COUNT(MC.nMovNro) as Numero , IsNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoDctoInactiva & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & I & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
             & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
             
        Sql3 = " SELECT COUNT(MC.nMovNro) as Numero , IsNull(Sum(MC.nMonto),0) as Monto FROM MOVCAP MC " _
             & " INNER JOIN MOV M On MC.nMovNro = M.nMOvNro WHERE MC.cOpeCod = '" & gAhoCancInact & "'" _
             & " And Substring(MC.cCtaCod,9,1) = '" & I & "' And M.nMovFlag <> " & gMovFlagExtornado & " " _
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
        
    lsCad = lsCad & Space(10) & "Moneda : " & IIf(I = 1, "SOLES", "DOLARES") & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Descripción                   " & Space(2) & "Nro     " & Space(2) & "     Monto     " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Paso de Activas a Inactivas : " & Space(2) & JDNum(Str(lnNum1), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon1), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Descuento de Inactivas      : " & Space(2) & JDNum(Str(lnNum2), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon2), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Cancelación de Inactivas    : " & Space(2) & JDNum(Str(lnNum3), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon3), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "                      TOTAL : " & Space(2) & JDNum(Str(lnTotNumero), 6, False, 6, 0) & Space(2) & JDNum(Str(lnTotMonto), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Next
    ConsolidadoInactivas = lsCad
End Function

Public Function RepCtaApe(pdFecIni As Date, pdFecFin As Date) As String
    Dim lsCad As String
    Dim lsFechas As String
    
    lsCad = ""
    lsFechas = " entre " & Format(pdFecIni, gsFormatoFechaView) & " y " & Format(pdFecFin, gsFormatoFechaView) & " "
    
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro Aperturadas SOLES" & lsFechas & "- Con Orden de Pago -", "1", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro Aperturadas SOLES" & lsFechas & "- Sin Orden de Pago -", "1", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Plazo Fijo SOLES" & lsFechas, "1", Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaApeAho("Cuentas de CTS SOLES" & lsFechas, "1", Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro DOLARES" & lsFechas & "- Con Orden de Pago -", "2", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Ahorro DOLARES" & lsFechas & "- Sin Orden de Pago -", "2", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaApeAho("Cuentas de Plazo Fijo DOLARES" & lsFechas, "2", Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaApeAho("Cuentas de CTS DOLARES" & lsFechas, "2", Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    RepCtaApe = lsCad
End Function

Private Function GetCtaApeAho(psTit As String, psMoneda As String, psProd As String, lnPagina As Long, lnItem As Long, pdFecIni As Date, pdFecFin As Date, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 35
    Dim lnMonto  As Currency
    Dim lnMontoApe  As Currency
    Dim lsFecha As String * 24
    Dim lsTipPers As String * 25
    Dim lsContador As String * 4
    
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
        lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;25; ;15;Monto;16; ;2;Fecha de Apertura;20; ;4;Personeria;25;Monto_Ape;18;", lnItem)
        
        lsCodCtaAnt = ""
        lnContador = 0
        
        While Not rsAho.EOF
            
            If rsAho!Cuenta <> lsCodCtaAnt Then
                lnContador = lnContador + 1
                lsContador = FillNum(Str(lnContador), 4, "0")
            Else
                lsContador = ""
            End If
                
            lsCodCta = IIf(rsAho!Cuenta = lsCodCtaAnt, "", rsAho!Cuenta)
            
            lsNomPers = PstaNombre(rsAho!Nombre)
            lnMonto = rsAho!Monto
            lnMontoApe = rsAho!MontoApe
            lsFecha = Format(rsAho!FechaApe, "dd/mm/yyyy hh:mm:ss AMPM")
            lsTipPers = rsAho!Tipo
            
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(2) & Space(18 - Len(Format(lnMonto, "#,##0.00"))) & Format(lnMonto, "#,##0.00") & Space(2) & lsFecha _
                     & lsTipPers & Space(18 - Len(Format(lnMontoApe, "#,##0.00"))) & Format(lnMontoApe, "#,##0.00") & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;25; ;25;Monto;16; ;2;Fecha de Apertura;20; ;4;Personeria;25;Monto_Ape;18;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If
    
    GetCtaApeAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function

Private Function DevQry(psCodProd As String, dFecIni As Date, dFecFin As Date, Optional psOrdPag As String = "S", Optional psMoneda As String = "1", Optional pbInactivas As Boolean = False) As String
'    If pbInactivas Then
'        DevQry = " Select Cuenta = A.cCodCta, Nombre = P.cNomPers, Monto = A.nSaldCntAC, FechaApe = A.dAperAC, " _
'               & " UltCon = dUltCntAC FROM Ahorroc A Inner Join Perscuenta PC" _
'               & " Inner Join " & gcCentralPers & "Persona P On P.cCodPers = PC.cCodPers " _
'               & " on A.cCodcTA = pC.cCodCta" _
'               & " where bInactiva = 1 and cOrdPag = '" & psOrdPag & "' and substring(Aho.cCodCta,6,1) = '" & psMoneda & "' and cEstCtaAC not in ('C','U') order by Aho.cCodCta"
'        Exit Function
'    End If
'
'    Select Case psCodProd
'        Case producto.gCapAhorros
'            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
'                   & " Case nPersoneria" _
'                   & " When 0 then 'PERSONA NATURAL'" _
'                   & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
'                   & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
'                   & " When 3 then 'CMAC LOCAL' End Tipo," _
'                   & " IsNull(nMonto,0) MontoApe" _
'                   & " From Producto PR" _
'                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
'                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
'                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
'                   & " Inner Join CaptacAhorros CAPAHO On PR.cCtaCod = CAPAHO.cCtaCod" _
'                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & gAhoApeEfec & "','" & gAhoApeNA & "','" & gAhoApeChq & "')" _
'                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "' and CAPAHO.bOrdPag = " & psOrdPag & "" _
'                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "'" _
'                   & " Order By CAP.cCtaCod"
'        Case Producto.gCapPlazoFijo
'            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
'                   & " Case nPersoneria" _
'                   & " When 0 then 'PERSONA NATURAL'" _
'                   & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
'                   & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
'                   & " When 3 then 'CMAC LOCAL' End Tipo," _
'                   & " IsNull(nApertura,0) MontoApe" _
'                   & " From Producto PR" _
'                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
'                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
'                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
'                   & " Inner Join CaptacPlazoFijo CAPPF On PR.cCtaCod = CAPPF.cCtaCod" _
'                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & gPFApeEfec & "','" & gPFApeNA & "','" & gPFApeChq & "')" _
'                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "'" _
'                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "'" _
'                   & " Order By CAP.cCtaCod"
'        Case producto.gCapCTS
'            DevQry = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto, CAP.dApertura FechaApe," _
'                   & " Case nPersoneria" _
'                   & " When 0 then 'PERSONA NATURAL'" _
'                   & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
'                   & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
'                   & " When 3 then 'CMAC LOCAL' End Tipo," _
'                   & " IsNull(nMonto,0) MontoApe" _
'                   & " From Producto PR" _
'                   & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
'                   & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
'                   & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
'                   & " Inner Join CaptacCTS CAPCTS On PR.cCtaCod = CAPCTS.cCtaCod" _
'                   & " Left Join MovCap MC On MC.cCtaCod = CAP.cCtaCod And MC.cOpeCod In ('" & gCTSApeEfec & "','" & gCTSApeNA & "','" & gCTSApeChq & "')" _
'                   & " Where CAP.dApertura Between '" & Format(dFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, dFecFin), gsFormatoFecha) & "'" _
'                   & " And SubString(PR.cCtaCod,9,1) = '" & psMoneda & "' and Left(nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "'" _
'                   & " Order By CAP.cCtaCod"
'    End Select
End Function

Public Function RepCtaCan(pdFecIni As Date, pdFecFin As Date) As String
    Dim lsCad As String
    Dim lsFechas As String
    
    lsCad = ""
    lsFechas = " entre " & Format(pdFecIni, gsFormatoFechaView) & " y " & Format(pdFecFin, gsFormatoFechaView) & " "
    
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro SOLES" & lsFechas & "- Con Orden de Pago -", "1", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro SOLES" & lsFechas & "- Sin Orden de Pago -", "1", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Plazo Fijo SOLES" & lsFechas & "", "1", Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de CTS SOLES" & lsFechas & "", "1", Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro DOLARES" & lsFechas & "- Con Orden de Pago -", "2", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "1")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Ahorro DOLARES" & lsFechas & "- Sin Orden de Pago -", "2", Producto.gCapAhorros, 0, 0, pdFecIni, pdFecFin, "0")
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de Plazo Fijo DOLARES" & lsFechas & "", "2", Producto.gCapPlazoFijo, 0, 0, pdFecIni, pdFecFin)
    lsCad = lsCad & GetCtaCanAho("Cuentas Canceladas de CTS DOLARES" & lsFechas & "", "2", Producto.gCapCTS, 0, 0, pdFecIni, pdFecFin)
    
    RepCtaCan = lsCad
End Function

Private Function GetCtaCanAho(psTit As String, psMoneda As String, psProd As String, lnPagina As Long, lnItem As Long, pdFecIni As Date, pdFecFin As Date, Optional psOrdPag As String = "S", Optional pbInactivas As Boolean = False) As String
    Dim sqlAho As String
    Dim rsAho As ADODB.Recordset
    Set rsAho = New ADODB.Recordset
    Dim lsCodCtaAnt As String
    
    Dim lsCodCta As String * 20
    Dim lsNomPers As String * 35
    Dim lnMonto  As Currency
    Dim lsFecha As String * 24
    Dim lsTipPers As String * 25
    Dim lsContador As String * 4
    Dim lsCadena As String
    Dim lnContador As Integer
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    lsCadena = ""
    
    sqlAho = DevQryCan(psProd, pdFecIni, pdFecFin, psOrdPag, psMoneda)
    Set rsAho = oCon.CargaRecordSet(sqlAho)
             
    If Not (rsAho.EOF And rsAho.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
        lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;17; ;23;Monto C.;16; ;2;Fecha de Ult.Mov;16; ;8;Personeria;15; ;10;", lnItem)
        
        lsCodCtaAnt = ""
        lnContador = 0
        
        While Not rsAho.EOF
            
            If rsAho!Cuenta <> lsCodCtaAnt Then
                lnContador = lnContador + 1
                lsContador = FillNum(Str(lnContador), 4, "0")
            Else
                lsContador = ""
            End If
                
            lsCodCta = IIf(rsAho!Cuenta = lsCodCtaAnt, "", rsAho!Cuenta)
            
            lsNomPers = PstaNombre(rsAho!Nombre)
            lnMonto = rsAho!Monto
            lsFecha = Format(rsAho!FechaCancel, gsFormatoFechaHoraView)
            lsTipPers = rsAho!Tipo
            
            lsCadena = lsCadena & lsContador & Space(2) & lsCodCta & lsNomPers & Space(2) & Space(18 - Len(Format(lnMonto, "#,##0.00"))) & Format(lnMonto, "#,##0.00") & Space(2) & lsFecha & lsTipPers & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psTit, lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, psMoneda)
                lsCadena = lsCadena & Encabezado("Item;6;Nro Cuenta;12; ;8;Nombre;17; ;23;Monto C.;16; ;2;Fecha de Ult.Mov;16; ;8;Personeria;15; ;10;", lnItem)
                lnItem = 0
            End If
        
            lsCodCtaAnt = rsAho!Cuenta
            rsAho.MoveNext
        Wend
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If
    
    GetCtaCanAho = lsCadena
    
    rsAho.Close
    Set rsAho = Nothing
End Function

Private Function DevQryCan(psCodProd As String, dFecIni As Date, dFecFin As Date, Optional psOrdPag As String = "S", Optional psMoneda As String = "1") As String
    
    Select Case psCodProd
        Case Producto.gCapAhorros
            DevQryCan = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto," _
                      & " Convert(DateTime," _
                      & " Substring(CAP.cUltimaActualizacion,5,2)" _
                      & " + '/' + Left(CAP.cUltimaActualizacion,4)" _
                      & " + '/'  + Substring(CAP.cUltimaActualizacion,7,2)" _
                      & " + ' '  + Substring(CAP.cUltimaActualizacion,9,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,11,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,13,2)" _
                      & " ) FechaCancel," _
                      & " Case CAP.nPersoneria" _
                      & " When 0 then 'PERSONA NATURAL'" _
                      & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                      & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                      & " When 3 then 'CMAC'" _
                      & " End  Tipo" _
                      & " From Producto PR" _
                      & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                      & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                      & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                      & " Inner Join CaptacAhorros CAPAHO On PR.cCtaCod = CAPAHO.cCtaCod" _
                      & " Where  CAP.cUltimaActualizacion between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'" _
                      & " And bOrdPag = " & psOrdPag & " and substring(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
                      & " And Left(PR.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "' Order By PR.cCtaCod"
        Case Producto.gCapPlazoFijo
            DevQryCan = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto," _
                      & " Convert(DateTime," _
                      & " Substring(CAP.cUltimaActualizacion,5,2)" _
                      & " + '/' + Left(CAP.cUltimaActualizacion,4)" _
                      & " + '/'  + Substring(CAP.cUltimaActualizacion,7,2)" _
                      & " + ' '  + Substring(CAP.cUltimaActualizacion,9,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,11,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,13,2)" _
                      & " ) FechaCancel," _
                      & " Case CAP.nPersoneria" _
                      & " When 0 then 'PERSONA NATURAL'" _
                      & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                      & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                      & " When 3 then 'CMAC'" _
                      & " End  Tipo" _
                      & " From Producto PR" _
                      & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                      & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                      & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                      & " Inner Join CaptacPlazoFijo CAPPF On PR.cCtaCod = CAPPF.cCtaCod" _
                      & " Where  CAP.cUltimaActualizacion between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'" _
                      & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
                      & " And Left(PR.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "' Order By PR.cCtaCod"
        Case Producto.gCapCTS
            DevQryCan = " Select  PR.cCtaCod Cuenta, PE.cPersNombre Nombre, CAP.nSaldoDisp Monto," _
                      & " Convert(DateTime," _
                      & " Substring(CAP.cUltimaActualizacion,5,2)" _
                      & " + '/' + Left(CAP.cUltimaActualizacion,4)" _
                      & " + '/'  + Substring(CAP.cUltimaActualizacion,7,2)" _
                      & " + ' '  + Substring(CAP.cUltimaActualizacion,9,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,11,2)" _
                      & " + ':'  + Substring(CAP.cUltimaActualizacion,13,2)" _
                      & " ) FechaCancel," _
                      & " Case CAP.nPersoneria" _
                      & " When 0 then 'PERSONA NATURAL'" _
                      & " When 1 then 'PERSONA JURIDICA SIN FINES DE LUCRO'" _
                      & " When 2 then 'PERSONA JURIDICA CON FINES DE LUCRO'" _
                      & " When 3 then 'CMAC'" _
                      & " End  Tipo" _
                      & " From Producto PR" _
                      & " Inner Join ProductoPersona PP On PR.cCtaCod = PP.cCtaCod" _
                      & " Inner Join Persona PE On PP.cPersCod  = PE.cPersCod" _
                      & " Inner Join Captaciones CAP On PR.cCtaCod = CAP.cCtaCod" _
                      & " Inner Join CaptacCTS CAPCTS On PR.cCtaCod = CAPCTS.cCtaCod" _
                      & " Where  CAP.cUltimaActualizacion between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'" _
                      & " And Substring(PR.cCtaCod,9,1) = '" & psMoneda & "'" _
                      & " And Left(PR.nPrdEstado,2) In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "') And Left(PR.cCtaCod,5) = '" & lsAgeCod & "' Order By PR.cCtaCod"
    End Select
    
End Function

Private Function CuentasInactivas(Moneda As String, dFecIni As Date, dFecFin As Date) As String
    Dim sql As String
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
    
    sql = " SELECT ISNULL(CANT.cCtaCodAnt,'') cCtaCodAnt, CAPAHO.cCtaCod cCodCta, MC.nSaldoDisponible nSaldDispAC, CAPAHO.nSaldoAnterior nSaldAntAC, MC.nMonto nMonTran, CAPAHO.dUltContacto dUltCntAC" _
        & " FROM CaptacAhorros CAPAHO" _
        & " Inner Join MovCap MC On CAPAHO.cCtaCod = MC.cCtaCod" _
        & " Inner Join Mov M On MC.nMovNro = M.nMovNro" _
        & " Left  Join CaptacCtaAntRel CANT On  CAPAHO.cCtaCod = CANT.cCtaCod" _
        & " WHERE MC.cOpeCod in ('" & gAhoCancInact & "','" & gAhoEstActInac & "') And Substring(CAPAHO.cCtaCod,9,1) = '" & Moneda & "'" _
        & " And M.nMovFlag <> " & gMovFlagExtornado & "" _
        & " And M.cMovNro Between '" & Format(dFecIni, gsFormatoMovFecha) & "' and  '" & Format(DateAdd("d", 1, dFecFin), gsFormatoMovFecha) & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
     vPag = 1
     lsCad = lsCad + oImpresora.gPrnSaltoLinea
     lsCad = lsCad + gsNomAge + Space(5) + "Sección Ahorro" + Space(5) + "Trujillo " + ArmaFecha(gdFecSis) + Space(4) + "Pagina: " + Format(vPag, "####") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
     lsCad = lsCad + Space(20) + "LISTADO DE CUENTAS INACTIVAS EN " + IIf(Moneda = "1", "SOLES", "DOLARES") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
     lsCad = lsCad + "CUENTA " + Space(5) + "CUENTAANT" + Space(5) + "S.ANTERIOR" + Space(5) + "S.DISPONIBLE" + Space(3) + "R.CAPITAL" + Space(2) + "F.ULTCONTACTO" + Space(5) + oImpresora.gPrnSaltoLinea
     lsCad = lsCad + String(90, "=") + oImpresora.gPrnSaltoLinea
     Do While Not rs.EOF
        CuentaA = rs!cCtaCodAnt
        RCapital = Abs(rs!nMonto)
        lsCad = lsCad + rs!cCodCta + Space(2) + CuentaA + JDNum(rs!nsaldantac, 13, True, 11, 2) + Space(2)
        lsCad = lsCad + JDNum(rs!nsalddispac, 13, True, 11, 2) + Space(4) + JDNum(Str(RCapital), 6, True, 4, 2) + Space(8) + Format(rs!dUltCntAC, gsFormatoFechaView) + oImpresora.gPrnSaltoLinea
        TDesc = TDesc + RCapital
        rs.MoveNext
        vLineas = vLineas + 1
        If vLineas > 52 Then
            vPag = vPag + 1
            lsCad = lsCad + oImpresora.gPrnSaltoLinea + "Resumen :" + Space(38) + JDNum(Str(TDesc), 13, True, 11, 2) + oImpresora.gPrnSaltoLinea
            lsCad = lsCad + oImpresora.gPrnSaltoPagina
            lsCad = lsCad + oImpresora.gPrnSaltoLinea
            lsCad = lsCad + gsNomAge + Space(5) + "Sección Ahorro" + Space(5) + "Trujillo " + ArmaFecha(gdFecSis) + Space(4) + "Pagina: " + Format(vPag, "####") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
            lsCad = lsCad + Space(20) + "LISTADO DE CUENTAS INACTIVAS EN " + IIf(Moneda = "1", "SOLES", "DOLARES") + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
            lsCad = lsCad + "CUENTA " + Space(5) + "CUENTAANT" + Space(5) + "S.ANTERIOR" + Space(5) + "S.DISPONIBLE" + Space(3) + "R.CAPITAL" + Space(2) + "F.ULTCONTACTO" + Space(5) + oImpresora.gPrnSaltoLinea
            lsCad = lsCad + String(90, "=") + oImpresora.gPrnSaltoLinea
            vLineas = 1
        End If
     Loop
       lsCad = lsCad + oImpresora.gPrnSaltoLinea + "Total Resumen :" + Space(32) + JDNum(Str(TDesc), 13, True, 11, 2) + oImpresora.gPrnSaltoLinea
    End If
    rs.Close
    Set rs = Nothing
    CuentasInactivas = IIf(lsCad <> "", lsCad & oImpresora.gPrnSaltoPagina, "")
End Function


