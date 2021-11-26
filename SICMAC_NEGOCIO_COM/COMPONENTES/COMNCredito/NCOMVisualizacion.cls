VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMVisualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim EspHorxPag As Integer

Private lsNegritaOn As String
Private lsNegritaOff  As String
Private lsSaltoLin As String
Private lsTab As String
Private EspVerxPag As Integer
Private nPuntPag As Integer

Dim oImpre As COMFunciones.FCOMImpresion
Dim oCad As COMFunciones.FCOMCadenas

Public Function ObtenerTitular(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select P.cPersNombre"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "' and PP.nPrdPersRelac=20"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        ObtenerTitular = rs!cPersNombre
    End If
    Set rs = Nothing
    
    Set oConec = Nothing
    
End Function

Public Function DatosGenerales(ByVal psCtaCod As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = " Select PT.nTasaInteres,CE.nPlazo,CE.nPeriodoFechaFija,CN.cConsDescripcion"
    sSql = sSql & " From ProductoTasaInteres PT"
    sSql = sSql & " Inner Join ColocacEstado CE on CE.cCtaCod=PT.cCtaCod and CE.nPrdEstado=2002"
    sSql = sSql & "Inner Join ColocacCred CC on CC.cCtaCod=CE.cCtaCod"
    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=CC.nColocCalendCod and CN.nConsValor<>3005 and CN.nConsCod=3005"
    sSql = sSql & " Where PT.cCtaCod='" & psCtaCod & " ' and PT.nPrdTasaInteres=1"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set DatosGenerales = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerCalendario(ByVal psCtaCod As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    ' calendario cuando el desembolso es el mismo dia de la aprobacion
    
    sSql = "Select CC.nCuota,Tipo=Case CC.nColocCalendApl"
    sSql = sSql & " When 1 Then 'Pago'"
    sSql = sSql & " When 0 Then 'Desembolso'"
    sSql = sSql & " End,"
    sSql = sSql & " CC.dVenc,"
    sSql = sSql & " Capital=isnull((Select nMonto From ColocCalendDet CD Where  CC.cCtaCod=CD.cCtaCod and CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
    sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1000),0),"
    sSql = sSql & " Interes=isnull((Select nMonto From ColocCalendDet CD Where  CC.cCtaCod=CD.cCtaCod and CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
    sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1100),0)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "' and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod=CC.cCtaCod)"
    sSql = sSql & " Order By CC.nCuota,CC.nColocCalendApl"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set VerCalendario = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function RelacionDesembolso(ByVal psCtaCod As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select P.cPersNombre,CN.cConsDescripcion as cRelacion"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=PP.nPrdPersRelac and CN.nConsCod=3002 and CN.nConsValor<>3002"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "'"
    sSql = sSql & " Order By PP.nPrdPersRelac"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RelacionDesembolso = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Private Function DameTipoCuota(ByVal pnTipoCuota As Integer) As Integer
        If pnTipoCuota = gColocCalendCodFFCC Or pnTipoCuota = gColocCalendCodFFCCPG Or pnTipoCuota = gColocCalendCodPFCC Or pnTipoCuota = gColocCalendCodPFCCPG Then
            DameTipoCuota = 2
        End If
        If pnTipoCuota = gColocCalendCodFFCF Or pnTipoCuota = gColocCalendCodFFCFPG Or pnTipoCuota = gColocCalendCodPFCF Or pnTipoCuota = gColocCalendCodPFCFPG Then
            DameTipoCuota = 1
        End If
        If pnTipoCuota = gColocCalendCodFFCD Or pnTipoCuota = gColocCalendCodFFCDPG Or pnTipoCuota = gColocCalendCodPFCD Or pnTipoCuota = gColocCalendCodPFCDPG Then
            DameTipoCuota = 3
        End If
End Function

Private Function DameTipoPeriodo(ByVal pnTipoPeriodo As Integer) As Integer
    If pnTipoPeriodo = gColocCalendCodFFCC Or pnTipoPeriodo = gColocCalendCodFFCCPG Or pnTipoPeriodo = gColocCalendCodFFCD Or pnTipoPeriodo = gColocCalendCodFFCCPG _
          Or pnTipoPeriodo = gColocCalendCodFFCCPG Or pnTipoPeriodo = gColocCalendCodFFCD Or pnTipoPeriodo = gColocCalendCodFFCDPG Or pnTipoPeriodo = gColocCalendCodFFCF Or pnTipoPeriodo = gColocCalendCodFFCFPG Then
            DameTipoPeriodo = 2
        End If
        If pnTipoPeriodo = gColocCalendCodPFCC Or pnTipoPeriodo = gColocCalendCodPFCCPG Or pnTipoPeriodo = gColocCalendCodPFCD Or pnTipoPeriodo = gColocCalendCodPFCCPG _
          Or pnTipoPeriodo = gColocCalendCodPFCCPG Or pnTipoPeriodo = gColocCalendCodPFCD Or pnTipoPeriodo = gColocCalendCodPFCDPG Or pnTipoPeriodo = gColocCalendCodPFCF Or pnTipoPeriodo = gColocCalendCodPFCFPG Then
            DameTipoPeriodo = 1
        End If
End Function

Public Function ChekingEqualsDateMoney(ByVal psCtaCod As String, ByVal dFecSist As Date) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select dPrdEstado From ColocacEstado"
    strSQL = strSQL & " Where nPrdEstado=2002 and cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(strSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If Format(rs!dPrdEstado, "dd/MM/yyyy") <> dFecSist Then
            ChekingEqualsDateMoney = True
        Else
            ChekingEqualsDateMoney = False
        End If
    End If
    Set rs = Nothing
End Function

Public Function ImpreDocumento(ByVal psTitular As String, ByVal pnTasaInteres As String, ByVal psPlazo As String, _
                          ByVal psTipoCuota As String, ByVal psCtaCod As String, _
                          ByVal dFecSis As Date, ByVal sCodUsu As String, ByVal rsPago As ADODB.Recordset) As String
    ' Considerando la estructura de la siguiente manera
    ' Cuota,tipo,dVenc,Capital,Interes
    
    Dim sCadImp As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim sNomAge As String
    
    Dim nCapital As Double
    Dim nInteres As Double
    
    ' Obteniendo el nombre de la Agencia
    sSql = "Select cAgeDescripcion From Agencias"
    sSql = sSql & " Where cAgeCod='" & Mid(psCtaCod, 4, 2) & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNomAge = rs!cAgeDescripcion
    End If
    
    Set rs = Nothing
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, sNomAge, dFecSis, sCodUsu, "PLAN DE PAGOS A DESEMBOLSAR", "CMAC ICA")
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lsTab & "Titular del Credito:" & psTitular & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tasa de Interes:" & pnTasaInteres & Chr$(10)
    sCadImp = sCadImp & lsTab & "Plazo del Credito:" & psPlazo & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tipo de Cuota:" & psTipoCuota
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
        
    sCadImp = sCadImp & "****************PLAN DE PAGOS***********************************" & Chr$(10)
    sCadImp = sCadImp & "Cuota" & Space(3) & "Tipo de Cuota" & Space(3) & "Fecha de Vencimiento" & Space(3) & "Capital" & Space(3) & "Interes" & Chr$(10)
    
    nCapital = 0
    nInteres = 0
    rsPago.MoveFirst
    Do Until rsPago.EOF
    If rsPago!tipo = "Desembolso" Then
       sCadImp = sCadImp & rsPago!nCuota & Space(7 - IIf(rsPago!nCuota > 9, 1, 0)) & rsPago!tipo & Space(6) & rsPago!dVenc & Space(13) & Format(rsPago!Capital, "#0.00") & Space(4) & Format(rsPago!Interes, "#0.00") & Chr$(10)
       nCapital = nCapital + rsPago!Capital
    Else
       sCadImp = sCadImp & rsPago!nCuota & Space(7 - IIf(rsPago!nCuota > 9, 1, 0)) & rsPago!tipo & Space(12) & rsPago!dVenc & Space(13) & rsPago!Capital & Space(5) & Format(rsPago!Interes, "#0.00") & Chr$(10)
       nInteres = nInteres + rsPago!Interes
    End If
        rsPago.MoveNext
    Loop
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & "Resumen de Montos a Pagar en " & IIf(Mid(psCtaCod, 9, 1) = "1", "Soles", "Dolares") & Chr$(10)
    sCadImp = sCadImp & "*********************************" & Chr$(10)
    sCadImp = sCadImp & "Capital total a pagar: " & Format(nCapital, "#0.00") & Chr$(10)
    sCadImp = sCadImp & "Interes total a pagar: " & Format(nInteres, "#0,00") & Chr$(10)
    
    ImpreDocumento = sCadImp
End Function

Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
                                ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "")
    
    
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "Fecha :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(45) & "Fecha :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
    End If
End Sub

Private Sub Class_Initialize()
    Set oImpre = New COMFunciones.FCOMImpresion
    Set oCad = New COMFunciones.FCOMCadenas
    EspHorxPag = 170
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
    lsSaltoLin = Chr$(10)
    EspHorxPag = 170
    EspVerxPag = 62
    nPuntPag = 0
    lsTab = Space(1)
End Sub

Public Function VerFteIngreso(ByVal pcNumFte As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select P.cPersNombre,FI.dPersFIinicio,FI.cRazSocDescrip,"
    sSql = sSql & " nSaldo = ((FID.nPersIngCli + FID.nPersIngCon) - (FID.nPersGastoFam + FID.nPersOtrIng))"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join PersFteIngreso FI on FI.cPersCod=P.cPersCod"
    sSql = sSql & " Inner Join PersFIDependiente FID on FID.cNumFuente=FI.cNumFuente"
    sSql = sSql & " Where FI.cNumFuente='" & pcNumFte & "'"
    sSql = sSql & " Union "
    sSql = sSql & " Select P.cPersNombre,FI.dPersFIinicio,FI.cRazSocDescrip,"
    sSql = sSql & " nSaldo = ((FII.nPersFIVentas + FII.nPersFIRecupCtasXCobrar + nPersIngFam) - (FII.nPersFICostoVentas + FII.nPersFIEgresosOtros + nPersEgrFam))"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join PersFteIngreso FI on FI.cPersCod=P.cPersCod"
    sSql = sSql & " Inner Join PersFIIndependiente FII on FII.cNumFuente=FI.cNumFuente"
    sSql = sSql & " Where FI.cNumFuente='" & pcNumFte & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set VerFteIngreso = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ImprimeComprobanteDesembolso(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
                                            ByVal psCodUsu As String, ByVal psNomCmac As String) As String

Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim oCred As COMDCredito.DCOMCredito
Dim oCalend As COMDCredito.DCOMCalendario
Dim R As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RDesemb As ADODB.Recordset
Dim sCadImp As String
Dim sPersNombreUser As String

    Set oCredDoc = New COMDCredito.DCOMCredDoc
    Set R = RecuperaDatosDesembolso(psCtaCod)
    sPersNombreUser = oCredDoc.ObtieneNombrePersonaUser(psCodUsu)

    Set oCredDoc = Nothing

    Set oCred = New COMDCredito.DCOMCredito
    Set REstadosCred = oCred.RecuperaEstadosdelCredito(psCtaCod, True, True, True)
    Set oCred = Nothing
    Set oCalend = New COMDCredito.DCOMCalendario

    Set oCalend = New COMDCredito.DCOMCalendario
    Set RDesemb = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing


    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "COMPROBANTE DE DESEMBOLSO", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito             : " & oImpre.ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Fecha de Aprobacion : " & oImpre.ImpreFormat(Format(R!dFecApr, "dd/mm/yyyy"), 20, 0) & "Codigo de Cliente : " & R!cPersCod & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & oImpre.ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & oCad.PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 20, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("CUOTAS", 18, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    REstadosCred.Close
    Set REstadosCred = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APODERADO         : " & oCad.PstaNombre(R!cApoderado) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA            : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO  : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INTERES           : " & Format(R!nTasaInteres, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TIPO DE DESEMBOLSO: " & R!cTipoDesemb & lnSaltoLinDoc

    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(25) & "DESEMBOLSOS EFECTUADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("F. PACTADA", 15, 0) & oImpre.ImpreFormat("FECHA", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("GASTOS", 15, 0) & oImpre.ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Format(RDesemb!dPago, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Trim(Str(RDesemb!nCuota)), 5, 0) & oImpre.ImpreFormat(RDesemb!nCapital, 15, 2) & oImpre.ImpreFormat(RDesemb!nGasto, 10, 2) & Space(10) & IIf(RDesemb!ncoloccalendestado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "________________________________   ______________________________    _________________________" & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & PstaNombre(R!cPersNombre) & Space(5) & PstaNombre(R!cAnalista) & Space(15) & "CAJA" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oCad.PstaNombre(R!cPersNombre) & Space(5) & oCad.PstaNombre(sPersNombreUser) & Space(19) & lnSaltoLinDoc
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    ImprimeComprobanteDesembolso = sCadImp

End Function

Public Function ObtenerNomAge(ByVal psCodAge As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select cAgeDescripcion"
    sSql = sSql & " From Agencias"
    sSql = sSql & " Where cAgeCod='" & psCodAge & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerNomAge = rs!cAgeDescripcion
    End If
    Set rs = Nothing
End Function


Function RecuperaDatosDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select PersAna.cPersNombre cAnalista, CE.dPrdEstado dFecApr, Pers.cPersCod,"
    sSql = sSql & " Pers.cPersNombre, CE.dPrdEstado dFecDesemb, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cMoneda,"
    sSql = sSql & " L.cDescripcion cLinea, P.nTasaInteres, CN2.cConsDescripcion cTipoDesemb, CN3.cConsDescripcion cTipoCred"
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 28"
    sSql = sSql & " Inner Join Persona PersAna ON PersAna.cPersCod = PP.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE ON P.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = 2002"
    sSql = sSql & " Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 20"
    sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP2.cPersCod"
    sSql = sSql & " Inner Join ProductoPersona PP3 ON P.cCtaCod = PP3.cCtaCod AND PP3.nPrdPersRelac = 29"
    sSql = sSql & " Inner Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
    sSql = sSql & " Inner Join Constante CN ON CN.nConsValor = Convert(Int, Substring(P.cCtaCod,9,1)) AND CN.nConsCod = 1011"
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtacod = P.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
    sSql = sSql & " Inner Join Constante CN2 ON CN2.nConsValor = CC.nTipoDesembolso AND CN2.nConsCod = 3018"
    sSql = sSql & " Left  Join Constante CN3 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = 1001"
    sSql = sSql & "Where P.cCtaCod = '" & psCtaCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaDatosDesembolso = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Private Function lnSaltoLinDoc() As String
    nPuntPag = nPuntPag + 1
    If nPuntPag > EspVerxPag Then
        nPuntPag = 0
        lnSaltoLinDoc = Chr$(12) & sCadCab
    Else
        lnSaltoLinDoc = Chr$(10)
    End If
    
End Function

'Public Function ImprimePlandePagos(ByVal psCtaCod As String, ByVal psNomAge As String, _
'    ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
'    ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
'    Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0) As String
'
'Dim R As ADODB.Recordset
'Dim oCredDoc As DCredDoc
'Dim oCalend As COMDCredito.DCOMCalendario
'Dim oNCred As NCOMCredito
'Dim MatCalend As Variant
'Dim sCadImp As String
'Dim i As Integer
'Dim nIntereses As Double
'Dim nCapital As Double
'Dim oParam As COMDCredito.DCOMParametro
'Dim nTramoNoConsPorcen As Double
'Dim RDesPar As ADODB.Recordset
'
'    Set oCredDoc = New DCredDoc
'    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
'    Set oCredDoc = Nothing
'
'    If pbMiViv Then
'        Set oParam = New COMDCredito.DCOMParametro
'        nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
'        Set oParam = Nothing
'        If Not pbCalendParalelo Then
'            pnMontoPrestamo = pnMontoPrestamo * ((100 - nTramoNoConsPorcen) / 100)
'            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
'        Else
'            pnMontoPrestamo = pnMontoPrestamo * (nTramoNoConsPorcen / 100)
'            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
'        End If
'    End If
'
'    Set oCalend = New COMDCredito.DCOMCalendario
'    Set RDesPar = oCalend.RecuperaCalendarioDesemb(psCtaCod)
'    Set oCalend = Nothing
'
'    Set oNCred = New NCOMCredito
'    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod, pnMontoPrestamo, pbCalendParalelo, IIf(pbCalendParalelo, True, False))
'    Set oNCred = Nothing
'
'    sCadImp = ""
'    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
'    sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & ocad.PstaNombre(R!cPersNombre) & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Analista      : " & ocad.PstaNombre(R!cAnalista) & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Interes       : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & "Gracia : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
'    If pbComodin Then
'        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Cuota Comodin", 30, 0)
'    ElseIf pbMiViv Then
'        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Mi Vivienda", 30, 0)
'    Else
'        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
'    End If
'
'    If pbCalendDin Then
'        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("SI", 30, 0)
'    Else
'        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)
'    End If
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & Chr$(10)
'
'    If RDesPar.RecordCount >= 2 Then
'        sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
'        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
'        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("DESEMBOLSO", 10, 0) & oImpre.ImpreFormat("ESTADO", 10, 2) & Chr$(10)
'        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
'        Do While Not RDesPar.EOF
'            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & oImpre.ImpreFormat(RDesPar!nCapital, 10, 2, True) & oImpre.ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
'            RDesPar.MoveNext
'        Loop
'    End If
'    RDesPar.Close
'
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 12, 0) & oImpre.ImpreFormat("Capital", 12, 0) & oImpre.ImpreFormat("Interes", 12, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gastos", 8, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
'    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
'    nIntereses = 0
'    nCapital = 0
'    For i = 0 To UBound(MatCalend) - 1
'        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(MatCalend(i, 0)), "ddd, d mmm yyyy"), 17, 0) & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 4, 0) & oImpre.ImpreFormat(CDbl(MatCalend(i, 3)) _
'                                                                                                                                                             + CDbl(MatCalend(i, 4)) _
'                                                                                                                                                             + CDbl(MatCalend(i, 5)) _
'                                                                                                                                                             + CDbl(MatCalend(i, 6)) _
'                                                                                                                                                             + CDbl(MatCalend(i, 7)) _
'                                                                                                                                                             + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)), 10, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 3)), 9, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 4)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 5)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 9)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 10, 2) & Chr$(10)
'        nIntereses = nIntereses + CDbl(MatCalend(i, 4))
'        nIntereses = CDbl(Format(nIntereses, "#0.00"))
'        nCapital = nCapital + CDbl(MatCalend(i, 3))
'        nCapital = CDbl(Format(nCapital, "#0.00"))
'        If i <> 0 And i Mod 44 = 0 Then
'            sCadImp = sCadImp & Chr$(12)
'            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", ImprimePlandePagos, lsTab, False)
'            sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & ocad.PstaNombre(R!cPersNombre) & Chr$(10) & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Analista      : " & ocad.PstaNombre(R!cAnalista) & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 20, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Interes       : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 20, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 20, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
'            sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & Chr$(10)
'            If pbComodin Then
'                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Cuota Comodin", 30, 0)
'            ElseIf pbMiViv Then
'                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Mi Vivienda", 30, 0)
'            Else
'                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
'            End If
'
'            If pbCalendDin Then
'                sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("SI", 30, 0)
'            Else
'                sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)
'            End If
'            sCadImp = sCadImp & Chr$(10)
'            sCadImp = sCadImp & Chr$(10)
'            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 12, 0) & oImpre.ImpreFormat("Capital", 12, 0) & oImpre.ImpreFormat("Interes", 12, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gastos", 8, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
'            sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
'        End If
'    Next i
'
'    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Totales : " & oImpre.ImpreFormat(nCapital, 32, 2) & oImpre.ImpreFormat(nIntereses, 7, 2) & oImpre.ImpreFormat(nCapital + nIntereses, 32, 2) & Chr$(10)
'    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
'    sCadImp = sCadImp & Chr$(27) & Chr$(107) & Chr$(2) & Chr$(27) & Chr$(83) & Chr$(1)
'    Set oNCred = New NCOMCredito
'    sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(oNCred.TasaIntPerDias(R!nTasaInteres, 360) * 100, "#0.00") & "%" & Chr$(10)
'    Set oNCred = Nothing
'    sCadImp = sCadImp & Chr$(27) & Chr$(84) & Chr$(1)
'    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(1)
'
'    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
'        sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
'        sCadImp = sCadImp & lsTab & "Si no se ha Realizado su Descuento por Planilla" & Chr$(10)
'        sCadImp = sCadImp & lsTab & "Acerquese a nuestras Agencias a Cancelar y Evite el cobro de Mora" & Chr$(10)
'    End If
'    If CInt(Mid(psCtaCod, 6, 3)) = gColHipoMiVivienda Or CInt(Mid(psCtaCod, 6, 3)) = gColHipoCaja Then
'        sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
'        sCadImp = sCadImp & lsTab & "Todas los Cuotas incluyen Gastos de Portes y Seguros." & Chr$(10)
'    End If
'    sCadImp = sCadImp & Chr$(12)
'    ImprimePlandePagos = sCadImp
'End Function

Public Function ObtenerMontoAprobacion(ByVal psCtaCod As String) As Double
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select nMonto"
    sSql = sSql & " From ColocacEstado"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "' And nPrdEstado=2002"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMontoAprobacion = rs!nMonto
    End If
    Set rs = Nothing
End Function


Public Function VerCreditoByPersona(ByVal psPersCod As String, Optional pbHistorial As Boolean = False, Optional ByVal pbTransferido As Boolean = False) As Recordset 'JUEZ 20121215
    'FRHU ERS022-2015 20150415 Se agrego pbTransferido
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select P.cCtaCod,CL.cDescripcion as cLineaCredDes,C.nMontoCol as nMontoDes,"
    sSql = sSql & " Moneda=Case When Substring(P.cCtaCod,9,1)='1' Then 'Soles' Else'Dolares' End,"
    sSql = sSql & " Pers1.cPersNombre as cPersNombre,CC.nDiasAtraso"
    sSql = sSql & " From Colocaciones C"
    sSql = sSql & " Inner Join ColocLineaCredito CL on CL.cLineaCred=C.cLineaCred"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=C.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Left  Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Inner Join Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    If pbHistorial Then 'JUEZ 20121215
        sSql = sSql & " Where Pers.cPersCod='" & psPersCod & "' and P.nPrdEstado in (2001,2002,2003,2020,2021,2022,2030,2031,2032,2201,2202,2203,2204,2205,2206,2050)"
    Else
        'FRHU 20150415 ERS022-2015
        'sSql = sSql & " Where Pers.cPersCod='" & psPersCod & "' and P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
        If pbTransferido Then
            sSql = sSql & " Where Pers.cPersCod='" & psPersCod & "' and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2301)"
        Else
            sSql = sSql & " Where Pers.cPersCod='" & psPersCod & "' and P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
        End If
        'FIN FRHU 20150415
    End If
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set VerCreditoByPersona = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerTitularByCredito(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select cPersNombre "
    sSql = sSql & " From Persona Pers"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerTitularByCredito = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
    End If
    Set rs = Nothing
End Function

Public Function CargarDatosVisualizacion(ByVal psCtaCod As String, _
                                        ByVal pdFecSis As Date, _
                                        ByRef psTitular As String, _
                                        ByRef prsDatos As ADODB.Recordset, _
                                        ByRef prsRelac As ADODB.Recordset, _
                                        ByRef prsCalend As ADODB.Recordset)
    
    Dim oCalend As COMNCredito.NCOMCalendario
    
    psTitular = ObtenerTitular(psCtaCod)
    
    Set prsDatos = DatosGenerales(psCtaCod)
    
    Set prsRelac = RelacionDesembolso(psCtaCod)

    If ChekingEqualsDateMoney(psCtaCod, pdFecSis) = True Then
        Set oCalend = New COMNCredito.NCOMCalendario
        Set prsCalend = oCalend.VerCalendarioD(psCtaCod, pdFecSis)
        Set oCalend = Nothing
    Else
        Set prsCalend = VerCalendario(psCtaCod)
    End If

End Function

Private Sub Class_Terminate()
    Set oImpre = Nothing
    Set oCad = Nothing
End Sub
'add jhcu
'Public Function VerSolWeb(ByVal gsCodAge As String, ByVal gsCodUser As String, Optional ByVal psNombre As String = "") As Recordset
Public Function VerSolWeb(ByVal gsCodAge As String, ByVal gsCodUser As String, Optional ByVal psNombre As String = "", Optional ByVal pnOpcion As Integer = -1) As Recordset 'Add JOEP20210206 garantia covid
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    'sSql = "EXEC stp_sel_sol_reprogramacion '" & gsCodAge & "','" & gsCodUser & "'," & "'" & psNombre & "'"
    sSql = "EXEC stp_sel_sol_reprogramacion '" & gsCodAge & "','" & gsCodUser & "'," & "'" & psNombre & "'," & pnOpcion & "" 'Add JOEP20210206 garantia covid
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set VerSolWeb = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
