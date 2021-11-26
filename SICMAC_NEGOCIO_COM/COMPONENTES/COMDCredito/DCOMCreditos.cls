VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum TipoExtorno
    gTipoExtornoCuenta = 0
    gTipoExtornoUsuario = 1
    gTipoExtornoCliente = 2
    gTipoExtornoGeneral = 3
End Enum
Dim vsServerCom As String 'madm 20100802
Dim vsServerPers As String 'madm 20100802
Dim csCentralCom As String 'RECO 20131012 ERS055:2013
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

'By Capi 28122007
Public Function ObtenerGarantiasInscritas(ByVal pdFecFinal As Date, ByVal psCodAge As String, ByVal psMonedas As String, ByVal pMatProd As Variant, ByVal sEstados As String) As ADODB.Recordset
    Dim sSql As String
    Dim sCadProd As String
    Dim sMonedas As String
    Dim i As Integer
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = Replace(Replace(psMonedas, "'", ""), " ", "")
    End If

    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = Replace(Replace(psMonedas, "'", ""), " ", "")
    End If
    
    ' BRGO 06/06/2010 BASILEA II
    'sSql = " Exec Stp_Sel_ObtenerGarantiasInscritas '" & Format(pdFecFinal, "yyyymmdd") & "','" & psCodAge & "','" & sCadProd & "','" & sMonedas & "','" & sEstados & "'"
    
    sSql = " Exec stp_sel_RS108386_GarantiasInscritas '" & Format(pdFecFinal, "yyyymmdd") & "','" & psCodAge & "','" & sCadProd & "','" & sMonedas & "','" & sEstados & "'"
    
    'FIN BRGO
    
    Set ObtenerGarantiasInscritas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function
 'By Capi 03112008
Public Function ReporteCreditosRefinanciados(ByVal pdFecFinal As Date, ByVal psCodAge As String, ByVal pnTipoCambio As Double) As ADODB.Recordset
    Dim sSql As String
    Dim sCadProd As String
    Dim sMonedas As String
    Dim i As Integer
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = " Exec Stp_Sel_ReporteCreditosRefinanciados'" & Format(pdFecFinal, "yyyymmdd") & "','" & psCodAge & "'," & pnTipoCambio
    Set ReporteCreditosRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function

'By Capi 28012008
Public Function ObtenerREULavadoDinero(ByVal pdFecInicial As String, ByVal pdFecFinal As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = " Exec Stp_Sel_ReporteREULavadoDeDinero '" & (pdFecInicial) & "','" & (pdFecFinal) & "','" & psCodAge & "'"
    Set ObtenerREULavadoDinero = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function
'By Capi 30012008
Public Function ObtenerDUDLavadoDinero(ByVal pdFecInicial As String, ByVal pdFecFinal As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = " Exec Stp_Sel_ReporteDudososLavadoDeDinero '" & (pdFecInicial) & "','" & (pdFecFinal) & "','" & psCodAge & "'"
    Set ObtenerDUDLavadoDinero = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function
Public Function RecuperaHistoriaCredPersona(ByVal psPersCod As String) As Variant
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim MatDatos() As String
Dim nPunt As Integer

    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    sSql = "Select TOP 3 P.cCtaCod, L.cDescripcion cLineaDesc, RH.cUser cAnalista, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, Lin.nTasaIni as nTasaGracia, "
    sSql = sSql & " CN2.cConsDescripcion cDestino, P.nTasaInteres, CN3.cConsDescripcion cTipoCuota, C.dVigencia,"
    sSql = sSql & " nNota = (select nColocNota  From ColocCalificacionAnalista Where cCtaCod = P.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = P.cCtaCod)), "
    sSql = sSql & " CE.nMonto nMontoSol, CE.dPrdEstado dFecSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
    sSql = sSql & " CE2.nMonto nMontoSug, CE2.dPrdEstado dFecSug, CE2.nCuotas nCuotasSug, CE2.nPlazo nPlazoSug, CE2.nPeriodoGracia nPeriodoGraciaSug,"
    sSql = sSql & " CE3.nMonto nMontoApr, CE3.dPrdEstado dFecApr, CE3.nCuotas nCuotasApr, CE3.nPlazo nPlazoApr, CE3.nPeriodoGracia nPeriodoGraciaApr, CN8.cConsDescripcion cTipoGracia,"
    sSql = sSql & " CN4.cConsDescripcion cMotivoRech, CE5.dPrdEstado dFecCancel, CE6.dPrdEstado dFecJud, CC.cMetLiquidacion, CC.cProtesto,"
    sSql = sSql & " CE7.nPrdEstado nEstRefin, CC.bCargoAuto, CN5.cConsDescripcion cEstActual, CN6.cConsDescripcion as cTipoCredDescrip, "
    sSql = sSql & " CN7.cConsDescripcion cTipoDesemb, CC.nNroProxDesemb, "
    sSql = sSql & " nCuotaSug = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=1) " & " ,"
    sSql = sSql & " nCuotaApr = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=2) "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdpersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSql = sSql & "                 Left Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
    sSql = sSql & "                 Left Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                 Left Join RRHH RH ON PP2.cPersCod = RH.cPersCod "
    sSql = sSql & "                 Left Join ProductoPersona PP3 ON PP3.cCtaCod = P.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSql = sSql & "                 Left Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = CC.nColocCondicion AND CN.nConsCod = " & gColocCredCondicion
    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCalendCod AND CN3.nConsCod = " & gColocTipoCalend
    sSql = sSql & "                 Left Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSql = sSql & "                 Left Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado = " & gColocEstSug
    sSql = sSql & "                 Left Join ColocacEstado CE3 ON CE3.cCtaCod = P.cCtaCod AND CE3.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                 Left Join ColocacEstado CE4 ON CE4.cCtaCod = P.cCtaCod AND CE4.nPrdEstado = " & gColocEstRech
    sSql = sSql & "                 Left Join Constante CN4 ON CN4.nConsValor = CE4.nMotivoRechazo AND CN4.nConsCod = " & gColocMotivRechazo
    sSql = sSql & "                 Left Join ColocacEstado CE5 ON CE5.cCtaCod = P.cCtaCod AND CE5.nPrdEstado = " & gColocEstCancelado
    sSql = sSql & "                 Left Join ColocacEstado CE6 ON CE6.cCtaCod = P.cCtaCod AND CE6.nPrdEstado = " & gColocEstJudicial
    sSql = sSql & "                 Left Join ColocacEstado CE7 ON CE7.cCtaCod = P.cCtaCod AND CE7.nPrdEstado = " & gColocEstRefNorm
    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = P.nPrdEstado AND CN5.nConsCod = " & gColocEstado
    sSql = sSql & "                 Left Join Constante CN6 ON convert(int,substring(P.cCtaCod,6,3)) = CN6.nConsValor AND CN6.nConsCod = " & gProducto
    sSql = sSql & "                 Left Join Constante CN7 ON CN7.nConsValor = CC.nTipoDesembolso AND CN7.nConsCod = " & gColocTiposDesembolso
    sSql = sSql & "                 Left Join Constante CN8 ON CN8.nConsValor = CE3.nTipoGracia AND CN8.nConsCod = " & gColocTiposGracia
    sSql = sSql & "                 Left Join ColocLineaCreditoTasa Lin ON C.cLineaCred = Lin.cLineaCred AND Lin.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntGracia
    sSql = sSql & " WHERE PP.cPersCod = '" & psPersCod & "' AND C.dVigencia is not NULL"
    sSql = sSql & " ORDER BY C.dVigencia DESC "
    Set R = oConn.CargaRecordSet(sSql)
    ReDim MatDatos(R.RecordCount, 9)
    nPunt = 0
    Do While Not R.EOF
        MatDatos(nPunt, 0) = R!cCtaCod
        MatDatos(nPunt, 1) = R!cEstActual
        MatDatos(nPunt, 2) = R!dVigencia
        MatDatos(nPunt, 3) = IIf(IsNull(R!dFecCancel), "", R!dFecCancel)
        If Not IsNull(R!nMontoApr) Then
            MatDatos(nPunt, 4) = Format(R!nMontoApr, "#0.00")
        Else
            If Not IsNull(R!nMontoSug) Then
                MatDatos(nPunt, 4) = Format(R!nMontoSug, "#0.00")
            Else
                MatDatos(nPunt, 4) = Format(R!nMontoSol, "#0.00")
            End If
        End If
        If Not IsNull(R!nCuotasApr) Then
            MatDatos(nPunt, 5) = R!nCuotasApr & "/" & R!nPlazoApr
        Else
            If Not IsNull(R!nCuotasSug) Then
                MatDatos(nPunt, 5) = R!nCuotasSug & "/" & R!nPlazoSug
            Else
                MatDatos(nPunt, 5) = R!nCuotasSol & "/" & R!nPlazoSol
            End If
        End If
        If Not IsNull(R!nCuotaApr) Then
            MatDatos(nPunt, 6) = Format(R!nCuotaApr, "#0.00")
        Else
            MatDatos(nPunt, 6) = Format(R!nCuotaSug, "#0.00")
        End If
        MatDatos(nPunt, 7) = Trim(Str(IIf(IsNull(R!nNota), 0, R!nNota)))
        MatDatos(nPunt, 8) = Trim(R!cAnalista)
        nPunt = nPunt + 1
        R.MoveNext
    Loop
    R.Close
    oConn.CierraConexion
    RecuperaHistoriaCredPersona = MatDatos
End Function

Public Function RecuperaPersonasEstadoCred(ByVal pnEstadoCred As Variant, Optional ByVal pMatProd As Variant = Nothing, _
    Optional ByVal pbRefin As Boolean = False, Optional pbMuestraTodos As Boolean = False) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim vsEstadoCred As String
Dim vsProductos As String
Dim i As Integer
Dim bSolicitado As Boolean
Dim bAprobado As Boolean

    On Error GoTo ErrorRecuperaPersonasEstadoCred
    vsEstadoCred = ""
    bSolicitado = False
    bAprobado = False
    For i = 0 To UBound(pnEstadoCred)
        vsEstadoCred = vsEstadoCred & pnEstadoCred(i) & ","
        If CInt(pnEstadoCred(i)) = gColocEstSolic Then
            bSolicitado = True
        End If
        If CInt(pnEstadoCred(i)) = gColocEstAprob Then
            bAprobado = True
        End If
    Next i
    vsEstadoCred = Mid(vsEstadoCred, 1, Len(vsEstadoCred) - 1)
    vsProductos = ""
    If IsArray(pMatProd) Then
        vsProductos = "('"
        For i = 0 To UBound(pMatProd)
            vsProductos = vsProductos & pMatProd(i) & "','"
        Next i
        vsProductos = Mid(vsProductos, 1, Len(vsProductos) - 2)
    End If
    vsProductos = vsProductos & ")"
    
    sSql = "Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
    sSql = sSql & " Inner join ColocacCred C ON PP.cCtaCod = C.cCtaCod "
    sSql = sSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
    sSql = sSql & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (" & Trim(vsEstadoCred) & ") "
    If Len(vsProductos) > 1 Then
        sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
    End If
    If bSolicitado And Not pbMuestraTodos Then
        If pbRefin Then
            sSql = sSql & " AND  C.cCtaCod in (Select cCtaCod From ColocacRefinanc )"
        Else
            sSql = sSql & " AND  C.cCtaCod not in (Select cCtaCod From ColocacRefinanc )"
        End If
    End If
    
    If bAprobado Then
        sSql = sSql & "  UNION ALL "
        sSql = sSql & "Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
        sSql = sSql & " Inner join ColocacCred C ON PP.cCtaCod = C.cCtaCod "
        sSql = sSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
        sSql = sSql & " Inner join ColocCalendario CCal ON CCal.cCtaCod = C.cCtaCod AND CCal.nCuota = C.nNroProxDesemb AND CCal.nColocCalendApl = 0 AND CCal.nNroCalen = C.nNroCalen AND CCal.nColocCalendEstado = 0 "
        sSql = sSql & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (2020,2021,2022)"
        
    End If
    sSql = sSql & " order by P.cPersNombre"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasEstadoCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaPersonasEstadoCred:
    
End Function

'--------
Public Function RecuperaPersonasEstadoCredAgencia(ByVal pnEstadoCred As Variant, ByVal psCodAgencia As String, _
    Optional ByVal pMatProd As Variant = Nothing, _
    Optional ByVal pbRefin As Boolean = False, _
    Optional pbMuestraTodos As Boolean = False, _
    Optional pbLeasing As Boolean = False, _
    Optional ByVal pbMicroMulti As Boolean = False, _
    Optional pbInfoGas As Boolean = False, _
    Optional pbAmpliado As Boolean = False, Optional pgsCodCargo As String = "") As ADODB.Recordset
    '*** BRGO 20111222 Se agregó el parámetro pbInfoGas
    'WIOR 20120518 Se agrego el parametro pbMicroMulti
    'LUCV20180417, Agregó pbAmpliado: según incidente en la edición de créditos ampliados
    'JOEP20190205 CP pgsCodCargo
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim vsEstadoCred As String
Dim vsProductos As String
Dim i As Integer
Dim bSolicitado As Boolean
Dim bAprobado As Boolean


    On Error GoTo ErrorRecuperaPersonasEstadoCred
    vsEstadoCred = ""
    bSolicitado = False
    bAprobado = False
    For i = 0 To UBound(pnEstadoCred)
        vsEstadoCred = vsEstadoCred & pnEstadoCred(i) & ","
        If CInt(pnEstadoCred(i)) = gColocEstSolic Then
            bSolicitado = True
        End If
        If CInt(pnEstadoCred(i)) = gColocEstAprob Then
            bAprobado = True
        End If
    Next i
    vsEstadoCred = Mid(vsEstadoCred, 1, Len(vsEstadoCred) - 1)
    vsProductos = ""
    If IsArray(pMatProd) Then
        vsProductos = "('"
        For i = 0 To UBound(pMatProd)
            vsProductos = vsProductos & pMatProd(i) & "','"
        Next i
        vsProductos = Mid(vsProductos, 1, Len(vsProductos) - 2)
    End If
    vsProductos = vsProductos & ")"
    
'    'Cometno JOEP20190302 CP
'    sSql = "Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
'    sSql = sSql & " Inner join ColocacCred C ON PP.cCtaCod = C.cCtaCod "
'    sSql = sSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
'    sSql = sSql & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (" & Trim(vsEstadoCred) & ") "
'    If psCodAgencia <> "ALL" Then
'        sSql = sSql & " AND SubString(Prd.cCtaCod,4,2)='" & psCodAgencia & "'"
'    End If
'
'    If Len(vsProductos) > 1 Then
'        sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
'    End If
'    If pbLeasing = False Then
'        If pbInfoGas = True Then 'BRGO 20111222
'            'sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) = '517'" 'ARLO20180905
'            sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) in (select valor from fn_getProdEquivTbl('517'))" 'ARLO20180905
'        Else
'            'sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) NOT IN ('515','516')" 'ARLO20180905
'            sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) NOT IN (select valor from fn_getProdEquivTbl('515,516'))" 'ARLO20180905
'        End If
'    ElseIf pbLeasing = True Then
'        'sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN ('515','516')" 'ARLO20180905
'        sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN (select valor from fn_getProdEquivTbl('515,516'))" 'ARLO20180905
'    End If
'    If bSolicitado And Not pbMuestraTodos Then
'        If pbRefin Then
'            sSql = sSql & " AND  C.cCtaCod in (Select cCtaCod From ColocacRefinanc )"
'        Else
'            sSql = sSql & " AND  C.cCtaCod not in (Select cCtaCod From ColocacRefinanc )"
'        End If
'
'        '->***** LUCV20180417
'        If pbAmpliado Then
'            sSql = sSql & " AND  C.cCtaCod in (Select cCtaCod From ColocacAmpliado )"
'        Else
'            sSql = sSql & " AND  C.cCtaCod not in (Select cCtaCod From ColocacAmpliado )"
'        End If
'        '<-***** Fin LUCV20180417
'    End If
'
'    If bAprobado Then
'    'LUCV20160512 - ERS004-2016(Vigencia de refinanciacion de creditos)
'     If pbRefin Then
'            sSql = sSql & " AND  C.cCtaCod in (Select cCtaCod From ColocacRefinanc )"
'         Else
'            sSql = sSql & " AND  C.cCtaCod not in (Select cCtaCod From ColocacRefinanc )"
'     End If
'    'Fin LUCV20160512
'
'        sSql = sSql & " UNION " 'WIOR 20150630 QUITO 'ALL'
'        sSql = sSql & " Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
'        sSql = sSql & " Inner join ColocacCred C ON PP.cCtaCod = C.cCtaCod "
'        sSql = sSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
'        sSql = sSql & " Inner join ColocCalendario CCal ON CCal.cCtaCod = C.cCtaCod AND CCal.nCuota = C.nNroProxDesemb AND CCal.nColocCalendApl = 0 AND CCal.nNroCalen = C.nNroCalen AND CCal.nColocCalendEstado = 0 "
'        sSql = sSql & " INNER JOIN ColocacEstado CE ON CE.cCtaCod =C.cCtaCod AND CE.nPrdEstado=Prd.nPrdEstado AND CE.nTipoDesembolso=1"
'        sSql = sSql & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (2020,2021,2022)"
'        If Len(vsProductos) > 1 Then
'            sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
'        End If
'        If pbLeasing = False Then
'            'sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) NOT IN ('515','516')" 'ARLO20180905
'            sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) NOT IN (select valor from fn_getProdEquivTbl('515,516'))" 'ARLO20180905
'        ElseIf pbLeasing = True Then
'            'sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN ('515','516')" 'ARLO20180905
'            sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN (select valor from fn_getProdEquivTbl('515,516'))" 'ARLO20180905
'        End If
'
'    End If
'
'     'JOEP20190205 CP
'    If pgsCodCargo = "006024" Or pgsCodCargo = "007026" Then
'        sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN ('703')"
'    Else
'        sSql = sSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) NOT IN ('703')"
'    End If
'    'JOEP20190205 CP
'
'    'WIOR 20120518 ********************************************************************
'    If pbMicroMulti Then
'        sSql = sSql & " AND (Prd.cCtaCod  IN (SELECT cCtaCod from ColocacMicroseguro) "
'        sSql = sSql & " OR Prd.cCtaCod  IN(SELECT cCtaCod from ColocacMultiriesgo)) "
'    End If
'    'WIOR FIN *************************************************************************
'
'    sSql = sSql & " order by P.cPersNombre"
'    'Cometno JOEP20190302 CP
    
    sSql = "exec stp_sel_RecuperaPersonasEstadoCredAgecia '" & Trim(vsEstadoCred) & "','" & psCodAgencia & "'," & Len(vsProductos) & "," & IIf(pbLeasing, 1, 0) & "," & IIf(pbInfoGas, 1, 0) & "," & IIf(bSolicitado, 1, 0) & "," & IIf(pbMuestraTodos, 1, 0) & "," & IIf(pbRefin, 1, 0) & "," & IIf(pbAmpliado, 1, 0) & " ," & IIf(pbMicroMulti, 1, 0) & " ," & IIf(bAprobado, 1, 0) & ",'" & pgsCodCargo & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasEstadoCredAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaPersonasEstadoCred:
    
End Function

Public Sub ActualizaNivel(ByVal nNivel As Integer, ByVal cDescrip As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizaNivel
    sSql = "UPDATE ColocCredNivelesTipos SET cDescripcion = '" & cDescrip & "'"
    sSql = sSql & " WHERE nNivel = " & Trim(Str(nNivel))
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizaNivel:
    Err.Raise Err.Number, "Nuevo Nivel", Err.Description
End Sub
Public Sub Nuevonivel(ByVal nNivel As Integer, ByVal cDescrip As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorNuevonivel
    sSql = "INSERT INTO ColocCredNivelesTipos(nNivel,cDescripcion)"
    sSql = sSql & " VALUES(" & Trim(Str(nNivel)) & ",'" & cDescrip & "')"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorNuevonivel:
    Err.Raise Err.Number, "Nuevo Nivel", Err.Description

End Sub
Public Sub EliminaNivel(ByVal nNivel As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaNivel
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "DELETE ColocCredNivelesTipos WHERE nNivel = " & Trim(Str(nNivel))
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminaNivel:
    Err.Raise Err.Number, "Elimina Nivel", Err.Description
End Sub
Public Function RecuperaNivAprPersona(ByVal sPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaNivAprPersona
    sSql = "Select cCodNiv, cPersCod from ColocCredPersNivelesApr Where cPersCod = '" & sPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Set RecuperaNivAprPersona = oConecta.CargaRecordSet(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaNivAprPersona:
    Err.Raise Err.Number, "Recupera Niveles de Analista", Err.Description
    
End Function
Public Function RecuperaNiveles() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaNiveles
    sSql = "Select nNivel,cDescripcion from ColocCredNivelesTipos"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNiveles = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaNiveles:
    Err.Raise Err.Number, "Recupera Niveles", Err.Description
End Function

Public Function RecuperaCreditosVigxInstitucion(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosxInstitucion
    
    sSql = "Select P.cCtaCod, P.nTasaInteres, "
    sSql = sSql & " nCuotaPend = (Select Min(nCuota) as nCuota From ColocCalendario Where cCtaCod = P.cCtaCod And nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & " and nNroCalen = CC.nNroCalen ),"
    sSql = sSql & " dFecVencPend = (select top 1 dVenc from ColocCalendario where cCtaCod = P.cCtaCod AND nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & " and nNroCalen = CC.nNroCalen  Order by nCuota ),"
    sSql = sSql & " nCuota = (Select Min(nCuota) From ColocCalendario Where cCtaCod = P.cCtaCod And nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & " and nNroCalen = CC.nNroCalen )"
    sSql = sSql & " from Producto P Inner Join ColocacConvenio PP "
    sSql = sSql & " ON P.cCtaCod = PP.cCtaCod "
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & " Where PP.cPersCod = '" & psPersCod & "' "
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigxInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCreditosxInstitucion:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function TieneNivelAprobAsignado(ByVal psCodNiv As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select * from ColocCredNivelesApr Where nNivel = " & Trim(psCodNiv)
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.BOF And R.EOF Then
        TieneNivelAprobAsignado = False
    Else
        TieneNivelAprobAsignado = True
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

Public Function ExisteNivelAprobacion(ByVal psCodNiv As String, ByVal psProduct As String, _
    ByVal psTpoCred As String, ByVal pnNivel As Integer, ByVal pnMoneda As Integer) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorExisteNivelAprobacion
    sSql = "Select * from ColocCredNivelesApr Where cCodNiv = '" & Trim(psCodNiv) & "' AND cProduct = '" & Trim(psProduct) & "' AND cTpoCred = '" & Trim(psTpoCred) & "' AND nNivel = " & pnNivel & " And nMoneda = " & pnMoneda
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.BOF And R.EOF Then
        ExisteNivelAprobacion = False
    Else
        ExisteNivelAprobacion = True
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteNivelAprobacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

Public Function RecuperaNivelesAprobacion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaNivelesAprobacion

    sSql = "Select substring(NA.cCodNiv,4,1) + substring(NA.cCodNiv,1,3) + substring(NA.cCodNiv,5,1) + substring(NA.cCodNiv,6,1) as cCodNiv "
    sSql = sSql & " , CP.cConsDescripcion + space(50) + NA.cProduct as cProducto, "
    sSql = sSql & " CT.cConsDescripcion + space(50) + NA.cTpoCred as cTipoCred,"
    sSql = sSql & " CT2.cConsDescripcion + space(50) + CONVERT(varchar(3),NA.nMoneda) as cMoneda,"
    sSql = sSql & " N.cDescripcion + space(50) + convert(varchar(5),N.nNivel) as cNivel, "
    sSql = sSql & " NA.nMontoMin , NA.nMontoMax "
    sSql = sSql & " from ColocCredNivelesApr NA inner join Constante CP ON convert(int,NA.cProduct) = CP.nConsValor AND convert(varchar(15),CP.nConsValor) not like '23_' AND CP.nConsValor <> 305 AND CP.nConsValor <> CP.nConsCod AND CP.nConsCod = 1001"
    sSql = sSql & " inner join Constante CT ON convert(int,NA.cTpoCred) = CT.nConsValor AND CT.nConsCod = 3014 AND CT.nConsValor <> CT.nConsCod "
    sSql = sSql & " inner join Constante CT2 ON convert(int,NA.nMoneda) = CT2.nConsValor AND CT2.nConsCod = 1011 AND CT2.nConsValor <> CT2.nConsCod "
    sSql = sSql & " inner join ColocCredNivelesTipos N ON N.nNivel = NA.nNivel "
    sSql = sSql & " ORDER BY substring(NA.cCodNiv,4,1) + substring(NA.cCodNiv,1,3) + substring(NA.cCodNiv,5,1)"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivelesAprobacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaNivelesAprobacion:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion", Err.Description
End Function

Public Function DatosPosicionClienteAhorro(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteAhorro

    Set Conn = New COMConecta.DCOMConecta
    
    'By Capi 12112008 para modificar saldo disponible caj aut
    
'    sSql = " Select  P.cCtaCod,A.cAgeDescripcion, "
'    sSql = sSql & " nSaldoCont= CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE P.nSaldo END, "
'    sSql = sSql & " nSaldoDisp = CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE nSaldoDisp END, "
'    sSql = sSql & " C.dApertura, CN.cConsDescripcion as cTipoAho, "
'    sSql = sSql & " CN2.cConsDescripcion as cEstado, CN3.cConsDescripcion as cParticip,"
'    sSql = sSql & " CN4.cConsDescripcion as cMotivo, "
'    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' when '2' then 'DOLARES' end, cctacodant= case when R.CCTACODANT is null then '' else r.cctacodant end  "
'    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod and  nPrdPersRelac not in (14)"
'    sSql = sSql & " Inner Join Captaciones C ON C.cCtaCod = P.cCtaCod"
'    sSql = sSql & " left join Relcuentas R ON R.cCTACOD=P.CCTACOD   "
'    sSql = sSql & " Inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
'    sSql = sSql & " Left join ProductoBloqueos PB ON PB.cCtaCod = P.cCtaCod AND PB.cMovNro = (Select Top 1 cMovNro From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC)"
'    sSql = sSql & " LEFT Join Constante CN ON CONVERT(Int,SUBSTRING(P.cCtaCod,6,3))=CN.nConsValor AND CN.nConsCod = 1001"
'    sSql = sSql & " LEFT join Constante CN2 ON CN2.nConsValor = P.nPrdEstado AND CN2.nConsCod = 2001"
'    sSql = sSql & " LEFT join Constante CN3 ON CN3.nConsValor = PP.nPrdPersrelac AND CN3.nConsCod = 2005"
'    sSql = sSql & " LEFT join Constante CN4 ON CN4.nConsValor = PB.nBlqMotivo AND CN4.nConsCod = 2007 "
'    sSql = sSql & " WHERE PP.cpersCod = '" & psPersCod & "'"
'
    
'    sSql = " Select  Distinct P.cCtaCod,A.cAgeDescripcion, "
'    sSql = sSql & " nSaldoCont= CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE P.nSaldo END, "
'    sSql = sSql & " nSaldoDisp = Case   Substring(C.cCtaCod,6,3) " _
'                & "                     When    '232' Then Case nPrdEstado " _
'                & "                                             When 1400 Then 0 " _
'                & "                                             When 1300 Then 0 " _
'                & "                                             Else C.nSaldoDisp -Isnull(AHC.nRetencion,0) " _
'                & "                                        End " _
'                & "                     When    '234' Then Case nPrdEstado " _
'                & "                                                 When 1400 Then 0 " _
'                & "                                                 When 1300 Then 0 " _
'                & "                                                 Else C.nSaldoDisp -Isnull(CTS.nRetencion,0) " _
'                & "                                         End " _
'                & "                     Else " _
'                & "                             Case nPrdEstado " _
'                & "                                When 1400 Then 0 " _
'                & "                                When 1300 Then 0 " _
'                & "                                Else nSaldoDisp " _
'                & "                             End " _
'                & "                End, "
'    sSql = sSql & " C.dApertura, CN.cConsDescripcion as cTipoAho, "
'    sSql = sSql & " CN2.cConsDescripcion as cEstado, CN3.cConsDescripcion as cParticip,"
'    'sSql = sSql & " CN4.cConsDescripcion as cMotivo, "
'    sSql = sSql & " cMotivo = CASE WHEN P.nPrdEstado NOT IN (1100,1200) THEN '' ELSE CN4.cConsDescripcion END ," ' RIRO20141027
'    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' when '2' then 'DOLARES' end, "
'    sSql = sSql & " cctacodant= case when R.CCTACODANT is null then '' else r.cctacodant end, C.nBloqueoParcial  "
'    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod and  nPrdPersRelac not in (14)"
'    sSql = sSql & " Inner Join Captaciones C ON C.cCtaCod = P.cCtaCod"
'    sSql = sSql & " left join Relcuentas R ON R.cCTACOD=P.CCTACOD   "
'    sSql = sSql & " Inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
'    sSql = sSql & " Left join ProductoBloqueos PB ON PB.cCtaCod = P.cCtaCod AND PB.cMovNro = (Select Top 1 cMovNro From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC)"
'    sSql = sSql & "     And PB.nBlqMotivo = (Select Top 1 nBlqMotivo From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC)" 'BRGO 01/11/2010
'    sSql = sSql & " LEFT Join Constante CN ON CONVERT(Int,SUBSTRING(P.cCtaCod,6,3))=CN.nConsValor AND CN.nConsCod = 1001"
'    sSql = sSql & " LEFT join Constante CN2 ON CN2.nConsValor = P.nPrdEstado AND CN2.nConsCod = 2001"
'    sSql = sSql & " LEFT join Constante CN3 ON CN3.nConsValor = PP.nPrdPersrelac AND CN3.nConsCod = 2005"
'    sSql = sSql & " LEFT join Constante CN4 ON CN4.nConsValor = PB.nBlqMotivo AND CN4.nConsCod = 2007 "
'    sSql = sSql & " left Join CaptacAhorros AHC On AHC.cCtaCod=P.cCtaCod "
'    sSql = sSql & " left Join CaptacCTS CTS On AHC.cCtaCod=P.cCtaCod "
'    sSql = sSql & " WHERE PP.cpersCod = '" & psPersCod & "'"
    
     'ANDE 20170216
    sSql = "exec sp_sel_DatosPosicionClienteAhorro '" & psPersCod & "'"
    'FIN ANDE
    
    Conn.AbreConexion
    Set DatosPosicionClienteAhorro = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    
ErrorDatosPosicionClienteAhorro:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function


Public Function DatosPosicionClienteJudicial(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta

'***********************************************************
'***********************************************************
'INICIO ORCR-20140913*********
''''
'''''MAVM 20100610 BAS II
'''''    sSql = " Select P.cCtaCod, CR.dIngRecup, P.cCtaCod, A.cAgeDescripcion,"
'''''    sSql = sSql & " cTipoCred=Case substring(P.cCtaCod,6,1) "
'''''    sSql = sSql & "               when '1' then 'COMERCIAL'"
'''''    sSql = sSql & "               when '2' then 'MICROEMPRESA'"
'''''    sSql = sSql & "               when '3' then 'CONSUMO'"
'''''    sSql = sSql & "               when '4' then 'HIPOTECARIO'"
'''''    sSql = sSql & "               end,"
'''''    sSql = sSql & " P.nPrdestado, CT.cConsDescripcion cEstado,"
'''''    sSql = sSql & " CT2.cConsDescripcion as cParticip, P.nSaldo,"
'''''    sSql = sSql & " dFecCast = (Select top 1 dPrdEstado from ColocacEstado where cctacod = P.cCtaCod AND nPrdEstado = " & gColocEstRecVigCast & "),"
'''''    sSql = sSql & " cMovNro = (select top 1 cMovNro from Mov Where nMovnro = (select MAX(nMovnro) from MovCol Where cCtaCod = P.cCtaCod))"
'''''    sSql = sSql & " From Producto P  Inner join ColocRecup CR ON P.cCtaCod = CR.cCtaCod"
'''''    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
'''''    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod = " & gColocEstado
'''''    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod and PP.nPrdPersRelac=20"
'''''    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod = " & gColocRelacPers & " "
'''''    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "'"
''''
''''    sSql = " Select P.cCtaCod, CR.dIngRecup, P.cCtaCod, A.cAgeDescripcion,"
''''    sSql = sSql & " isnull(CN4.nConsValor,0) nTipoProdCod, ISNULL(CN4.cConsDescripcion,'') as cTipoProdDescrip, isnull(CN5.nConsValor,0) nTipoCredCod, ISNULL(CN5.cConsDescripcion,'') as cTipoCredDescrip,"
''''    sSql = sSql & " P.nPrdestado, CT.cConsDescripcion cEstado,"
''''    sSql = sSql & " CT2.cConsDescripcion as cParticip, P.nSaldo,"
''''    sSql = sSql & " dFecCast = (Select top 1 dPrdEstado from ColocacEstado where cctacod = P.cCtaCod AND nPrdEstado = " & gColocEstRecVigCast & "),"
''''    sSql = sSql & " cMovNro = (select top 1 cMovNro from Mov Where nMovnro = (select MAX(nMovnro) from MovCol Where cCtaCod = P.cCtaCod))"
''''    sSql = sSql & " From Producto P  Inner join ColocRecup CR ON P.cCtaCod = CR.cCtaCod"
''''    sSql = sSql & " inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod"
''''    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
''''    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod = " & gColocEstado
''''    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod and PP.nPrdPersRelac=20"
''''    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod = " & gColocRelacPers & " "
''''    sSql = sSql & " LEFT JOIN Constante CN4 ON CL.cTpoProdCod = CN4.nConsValor AND CN4.nConsCod = " & COMDConstantes.gTpoProducto
''''    sSql = sSql & " LEFT JOIN Constante CN5 ON CL.cTpoCredCod = CN5.nConsValor AND CN5.nConsCod = " & COMDConstantes.gTpoCredito
''''    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "'"
    
    
    sSql = "exec sp_sel_DatosPosicionClienteJudicial " & gColocEstRecVigCast & ", " & gColocEstado & ", " & gColocRelacPers & ", " & COMDConstantes.gTpoProducto & ", " _
    & COMDConstantes.gTpoCredito & ", '" & psPersCod & "'"
    
'***********************************************************
'***********************************************************
'FIN ORCR-20140913************
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set DatosPosicionClienteJudicial = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    

End Function


Public Function DatosPosicionClientePigno(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
'COMENTADO POR APRI
'    sSql = "Select P.cCtacod, P.nPrdEstado, CT.cConsDescripcion cEstado, C.dVigencia,"
'    sSql = sSql & " A.cAgeDescripcion, CP.nNroRenov, C.nMontoCol, C.dVenc, SUBSTRING(P.cCtaCod,4,2) as cAgeCod,"
'    sSql = sSql & " P.nSaldo , CT2.cConsDescripcion cParticip,(case when substring(p.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ) sMoneda ,"
'    sSql = sSql & " CP.NTASACION "
'    sSql = sSql & " From Producto P"
'    sSql = sSql & " Inner Join ColocPignoraticio CP  ON P.cCtaCod = CP.cCtaCod"
'    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
'    sSql = sSql & " Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod"
'    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
'    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod"
'    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod =" & gColocRelacPers
'    sSql = sSql & " Where PP.cPersCod = '" & psPersCod & "'"
    sSql = "EXEC stp_sel_DatosPosicionClientePigno '" & psPersCod & "'" 'APRI20180831 TIC1808290005
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set DatosPosicionClientePigno = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    

End Function

Public Function DatosPosicionClientePignoLima(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta

    sSql = "Select P.cCtacod, P.nPrdEstado, CT.cConsDescripcion cEstado, C.dVigencia,"
    sSql = sSql & " A.cAgeDescripcion, CP.nNroAmort nNroRenov, C.nMontoCol, C.dVenc, SUBSTRING(P.cCtaCod,4,2) as cAgeCod,"
    sSql = sSql & " P.nSaldo , CT2.cConsDescripcion cParticip "
    sSql = sSql & " From Producto P "
    sSql = sSql & " Inner Join ColocPigno CP  ON P.cCtaCod = CP.cCtaCod"
    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
    sSql = sSql & " Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod"
    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod"
    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod =" & gColocRelacPers
    sSql = sSql & " Where PP.cPersCod = '" & psPersCod & "'"
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set DatosPosicionClientePignoLima = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

End Function

'Public Function DatosPosicionCliente(ByVal psPersCod As String) As ADODB.Recordset
'Dim Conn As COMConecta.DCOMConecta
'Dim sSql As String
'
'On Error GoTo ErrorDatosPosicionCliente
'
'    Set Conn = New COMConecta.DCOMConecta
'    sSql = "Select P.cCtaCod,"
'    sSql = sSql & " dSolicitado=(SELECT CE.dPrdEstado From ColocacEstado CE where  CE.cCtaCod = P.cCtaCod and CE.nPrdEstado = " & gColocEstSolic & "),"
'    sSql = sSql & " A.cAgeDescripcion,"
'    sSql = sSql & " cTipoCred=Case substring(P.cCtaCod,6,1)"
'                              sSql = sSql & " when '1' then 'COMERCIAL'"
'                              sSql = sSql & " when '2' then 'MICROEMPRESA'"
'                              sSql = sSql & " when '3' then 'CONSUMO'"
'                              sSql = sSql & " when '4' then 'HIPOTECARIO'"
'                              sSql = sSql & " end,"
'    sSql = sSql & " P.nPrdEstado, CT.cConsDescripcion as cEstadoDesc, PP.nPrdPersRelac, CT2.cConsDescripcion as cRelacionDesc,"
'    sSql = sSql & " cPersAnalista = (select cUser From RRHH Where cPerscod = (Select top 1 cPersCod from ProductoPersona where nPrdPersRelac = '" & gColRelPersAnalista & "' and cCtaCod = P.cCtaCod)),"
'    sSql = sSql & " nAnalistaNota = (Select nColocNota from ColocCalificacionAnalista where cCtaCod = P.cCtaCod and dColocNotaFecha = (select Max(dColocNotaFecha) from  ColocCalificacionAnalista where cCtaCod = P.cCtaCod )),"
'    sSql = sSql & " CL.nMontoCol as nPrestamo, P.nSaldo, RC.cCodAnt1, RC.cCodAnt2,"
'    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1)"
'                    sSql = sSql & " when '1' then 'SOLES'"
'                    sSql = sSql & " when '2' then 'DOLARES'"
'                    sSql = sSql & " end,"
'    sSql = sSql & " dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & ")"
'    sSql = sSql & " from Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
'    sSql = sSql & " inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
'    sSql = sSql & " inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
'    sSql = sSql & " inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod"
'    sSql = sSql & " left join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
'    sSql = sSql & " left join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='" & gColocRelacPers & "'"
'    sSql = sSql & " left join RelCtaCred RC ON PP.cCtaCod = RC.cCodCta"
'    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "' AND SUBSTRING(P.cCtaCod,6,3) <> 121"
'
'    Conn.AbreConexion
'    Set DatosPosicionCliente = Conn.CargaRecordSet(sSql)
'    Conn.CierraConexion
'    Set Conn = Nothing
'    Exit Function
'
'ErrorDatosPosicionCliente:
'    Err.Raise Err.Number, "Credi//to Posicion Cliente", Err.Description
'End Function

Public Function DatosPosicionCliente(ByVal psPersCod As String, Optional ByVal bMuestraAnalistas As Boolean = True, Optional pdFecSis As Date = "1900/01/01", Optional nProm As Integer = 0) As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionCliente

    Set Conn = New COMConecta.DCOMConecta

'***********************************************************
'***********************************************************
'INICIO ORCR-20140913*********
''''    sSql = "Select P.cCtaCod,"
''''    sSql = sSql & " dSolicitado=(SELECT CE.dPrdEstado From ColocacEstado CE where  CE.cCtaCod = P.cCtaCod and CE.nPrdEstado = " & gColocEstSolic & "),"
''''    sSql = sSql & " A.cAgeDescripcion,"
''''    'MAVM 20100610 BAS II ***
''''    '    sSql = sSql & " cTipoCred=Case substring(P.cCtaCod,6,1)"
''''    '                              sSql = sSql & " when '1' then 'COMERCIAL'"
''''    '                              sSql = sSql & " when '2' then 'MICROEMPRESA'"
''''    '                              sSql = sSql & " when '3' then 'CONSUMO'"
''''    '                              sSql = sSql & " when '4' then 'HIPOTECARIO'"
''''    '                              sSql = sSql & " end,"
''''    sSql = sSql & " isnull(CN4.nConsValor,0) nTipoProdCod, ISNULL(CN4.cConsDescripcion,'') as cTipoProdDescrip,"
''''    sSql = sSql & " isnull(CN5.nConsValor,0) nTipoCredCod, ISNULL(CN5.cConsDescripcion,'') as cTipoCredDescrip,"
''''    '***
''''    sSql = sSql & " P.nPrdEstado, CT.cConsDescripcion as cEstadoDesc, PP.nPrdPersRelac, CT2.cConsDescripcion as cRelacionDesc,"
''''    sSql = sSql & " cPersAnalista = (select cUser From RRHH Where cPerscod = (Select top 1 cPersCod from ProductoPersona where nPrdPersRelac = '" & gColRelPersAnalista & "' and cCtaCod = P.cCtaCod)),"
''''
''''    'MAVM 20090914 ***
''''    sSql = sSql & " AnalistaInicial = ISNULL((Select Top 1 cUser from ColocacReasignaCartera C inner join RRHH R on C.cPersCodAnterior = R.cPersCod Where cCtaCod = P.cCtaCod Order By dReasignacion Asc), (select cUser From RRHH Where cPerscod = (Select top 1 cPersCod from ProductoPersona where nPrdPersRelac = '28' and cCtaCod = P.cCtaCod))),"
''''    'MAVM 20090914 ***
''''
''''    sSql = sSql & " nAnalistaNota = (Select nColocNota from ColocCalificacionAnalista where cCtaCod = P.cCtaCod and dColocNotaFecha = (select Max(dColocNotaFecha) from  ColocCalificacionAnalista where cCtaCod = P.cCtaCod )),"
''''    sSql = sSql & " CL.nMontoCol as nPrestamo, P.nSaldo, " 'RC.cCodAnt1, RC.cCodAnt2,"
''''    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1)"
''''                    sSql = sSql & " when '1' then 'SOLES'"
''''                    sSql = sSql & " when '2' then 'DOLARES'"
''''                    sSql = sSql & " end,"
''''    'sSql = sSql & " dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & "),"
''''    sSql = sSql & " dCancelado = (select dPrdEstado from Producto Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & "),"
''''    sSql = sSql & " dAprobacion=(Select  CC1.dPago"
''''        sSql = sSql & " From ColocCalendario CC1"
''''        sSql = sSql & " Where CC1.cCtaCod=P.cCtaCod and CC1.nColocCalendApl=0 and CC1.nCuota=1"
''''        sSql = sSql & " and CC1.nNroCalen=(Select Max(nNroCalen) From ColocCalendario"
''''        sSql = sSql & " Where cCtaCod=CC1.cCtaCod and CC1.nColocCalendApl=nColocCalendApl and"
''''        sSql = sSql & " nCuota=CC1.nCuota and CC1.nColocCalendEstado=nColocCalendEstado)"
''''        sSql = sSql & " and CC1.nColocCalendEstado=1),"
''''    '**DAOR 20071105************
''''    sSql = sSql & " dDesembolso = (Select  top 1 dPrdEstado From ColocacEstado Where cCtaCod=P.cCtaCod and nPrdEstado in (2020,2030) order by dPrdEstado), "
''''    '***************************
''''    '**DAOR 20071129************
''''    sSql = sSql & " CT3.cConsDescripcion as cCondicion, "
''''    '***************************
''''    'sSql = sSql & " cRFA=Case When CC.cRFA Is Null or CC.cRFA='NOR' Then 'NOR' Else CC.cRFA End"
''''    'ALPA****20080818****/******************************************************************
''''    If nProm = 0 Then
''''    sSql = sSql & " nDiasAtrasoAcum=ISNULL(nDiasAtrasoAcum,0)" 'CUSCO
''''    Else
''''        sSql = sSql & " nDiasAtrasoAcum=(select isnull(sum(case when isnull(datediff(day,ColC.dVenc,case when ColC.nColocCalendEstado=1 then ColC.dPago else '" & Format(pdFecSis, "YYYY/MM/DD") & "' end),0) <0 then 0 "
''''        sSql = sSql & " else isnull(datediff(day,ColC.dVenc,case when ColC.nColocCalendEstado=1 then ColC.dPago else '" & Format(pdFecSis, "YYYY/MM/DD") & "' end),0) end)/count(*),0)"
''''        sSql = sSql & " from ColocCalendario ColC inner Join Producto Pro on (ColC.cCtaCod=Pro.cCtaCod) "
''''        sSql = sSql & " where ColC.cCtaCod=P.cCtaCod and Pro.nPrdEstado not in (2001,2002,2003,2004) "
''''        sSql = sSql & " and ColC.nColocCalendApl= 1"
''''        sSql = sSql & " and ColC.nNroCalen = (select max(nNroCalen) from ColocCalendario where cCtaCod = CC.cCtaCod)"
''''        sSql = sSql & " and (ColC.nColocCalendEstado=1 "
''''        sSql = sSql & " or ColC.dVenc<='" & Format(pdFecSis, "YYYY/MM/DD") & "'))"
''''    End If
''''    'End ALPA****************/***************************************************************
''''    'ARCV 31-01-2007
''''    'sSql = sSql & " ,cCalif= dbo.ObtenerCalificacionDiasAtraso_Cuenta (P.cCtaCod)"
''''    'MAVM 20100826 Para los creditos estado retirado mostrar calificacion ""
''''    sSql = sSql & " ,Case CT.nConsValor When '2080' Then '' Else  dbo.ObtenerCalificacionDiasAtraso_Cuenta (P.cCtaCod) End cCalif "
''''    'ALPA****20080818**/********************************************************************
''''    sSql = sSql & ",isnull((select 1 as x from ProductoPersona PP2 inner join producto P2 on (PP2.cCtaCod=P2.cCtaCod)"
''''    'sSql = sSql & " where P2.nPrdEstado in (2032,2031,2030,2021,2022,2020) and"
''''    sSql = sSql & " where P2.nPrdEstado in (2001,2002,2003,2004) and"
''''    sSql = sSql & " PP2.nPrdPersRelac=PP.nPrdPersRelac and PP2.cPersCod=PP.cPersCod and PP2.cCtaCod =P.cCtaCod),0) as nNotVigenteCa, "
''''    sSql = sSql & "isnull((select count(PP2.cPersCod) from ProductoPersona PP2 inner join producto P2 on (PP2.cCtaCod=P2.cCtaCod)"
''''    sSql = sSql & " where P2.nPrdEstado in (2032,2031,2030,2021,2022,2020) and"
''''    sSql = sSql & " PP2.nPrdPersRelac=PP.nPrdPersRelac and PP2.cPersCod=PP.cPersCod and PP2.cCtaCod =P.cCtaCod),0) as nCont, "
''''    sSql = sSql & " /**/"
''''    sSql = sSql & " (select top 1 P2.cCtaCod  from ProductoPersona PP2 inner join producto P2 on (PP2.cCtaCod=P2.cCtaCod)"
''''    sSql = sSql & " where P2.nPrdEstado in (2050) and"
''''    sSql = sSql & " PP2.nPrdPersRelac = PP.nPrdPersRelac And PP2.cPersCod = PP.cPersCod"
''''    sSql = sSql & " and P2.dPrdEstado=(select max(P21.dPrdEstado) from ProductoPersona PP21 inner join producto P21 on (PP21.cCtaCod=P21.cCtaCod)"
''''    sSql = sSql & " where P21.nPrdEstado in (2050) and"
''''    sSql = sSql & " PP21.nPrdPersRelac=PP2.nPrdPersRelac and PP21.cPersCod=PP2.cPersCod)) as cCtaCodUL "
''''    sSql = sSql & " /**/"
''''    'End ALPA*******************************************************************************
''''    '----------------
''''    sSql = sSql & " from Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
''''    'sSql = sSql & " inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
''''    sSql = sSql & " inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
''''    sSql = sSql & " inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod"
''''    sSql = sSql & " inner join Agencias A ON  A.cAgeCod = CL.cAgeCodAct"
''''    sSql = sSql & " left join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
''''    sSql = sSql & " left join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='" & gColocRelacPers & "'"
''''    '**DAOR 20071129, Condición del crédito**********************///
''''    sSql = sSql & " left join Constante CT3 ON CC.nColocCondicion = CT3.nConsValor And CT3.nConsCod ='" & gColocCredCondicion & "'"
''''    '***********************************************************
'''''    sSQL = sSQL & " left join RelCtaCred RC ON PP.cCtaCod = RC.cCodCta"
''''    'MAVM 20100610 BAS II ***
''''    'sSql = sSql & " LEFT JOIN Constante CN4 ON convert(int,substring(cTpoProdCod,1,2)+'0') = CN4.nConsValor AND CN4.nConsCod = 3033 "
''''    'sSql = sSql & " LEFT JOIN Constante CN5 ON convert(int,substring(cTpoCredCod,1,2)+'0') = CN5.nConsValor AND CN5.nConsCod = 3034 "
''''    sSql = sSql & " LEFT JOIN Constante CN4 ON CL.cTpoProdCod = CN4.nConsValor AND CN4.nConsCod = " & COMDConstantes.gTpoProducto
''''    sSql = sSql & " LEFT JOIN Constante CN5 ON CL.cTpoCredCod = CN5.nConsValor AND CN5.nConsCod = " & COMDConstantes.gTpoCredito
''''    '***
''''    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "' AND SUBSTRING(P.cCtaCod,6,3) <> 121"
''''
''''    If nMuestraAnalistas = False Then
''''        sSql = sSql & " and nprdpersrelac <>" & gColRelPersAnalista & "  and nprdpersrelac <>" & gColRelPersApoderado
''''    End If
''''
''''    sSql = sSql & " Order By dSolicitado"
    
    '***********************************************************
    
    sSql = "exec sp_sel_DatosPosicionCliente " & gColocEstSolic & ", " & gColRelPersAnalista & ", " & gColocEstCancelado & ", '" & Format(pdFecSis, "YYYY/MM/DD") & "', " _
    & gColocEstado & ", " & gColocRelacPers & ", " & gColocCredCondicion & ", " & gTpoProducto & ", " _
    & gTpoCredito & ", '" & psPersCod & "', " & gColRelPersApoderado & ", " & nProm & ", " & CInt(bMuestraAnalistas)
    
    '    @gColocEstSolic int,
    '    @gColRelPersAnalista int,
    '    @gColocEstCancelado int,
    '    @pdFecSis varchar(10),
    
    '    @gColocEstado int,
    '    @gColocRelacPers int,
    '    @gColocCredCondicion int,
    '    @gTpoProducto int,
    
    '    @gTpoCredito int,
    '    @psPersCod varchar(13),
    '    @gColRelPersApoderado int,
    '    @nProm int,
    '    @nMuestraAnalistas bit
    
    '***********************************************************
    'FIN ORCR-20140913************
    
    Conn.AbreConexion
    Set DatosPosicionCliente = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    
ErrorDatosPosicionCliente:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Function DatosPosicionClienteCF(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteCF

    'peac 20071116 se agrego << and PP.nPrdPersRelac not in (28,29,36) >>
    Set Conn = New COMConecta.DCOMConecta
    'MAVM 20100612 BAS II
'    sSql = "Select P.cCtaCod, dEmitido =(SELECT top 1 CF.dPrdEstado From ColocCFEstado CF " _
'        & "WHERE CF.cCtaCod = P.cCtaCod and CF.nPrdEstado = 2020 order by CF.dPrdEstado), A.cAgeDescripcion, P.nPrdEstado, " _
'        & "CT.cConsDescripcion as cEstadoDesc, PP.nPrdPersRelac, CT2.cConsDescripcion as cRelacionDesc, " _
'        & "cPersAnalista = (select cUser From RRHH  " _
'        & "        Where cPerscod = (Select top 1 cPersCod from ProductoPersona WHERE nPrdPersRelac = '20' and cCtaCod = P.cCtaCod)), " _
'        & "CL.nMontoCol as nPrestamo, P.nSaldo, sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' " _
'        & "WHEN '2' then 'DOLARES' END, dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And " _
'        & "         nPrdEstado = 2090) " _
'        & " FROM Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac not in (28,29,36)" _
'        & " INNER JOIN Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod " _
'        & " LEFT JOIN Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =3001 " _
'        & " LEFT JOIN Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='3002' "
'    sSql = sSql & " LEFT JOIN ColocCartaFianza CCF ON PP.cCtaCod =CCF.cCtaCod " 'WIOR 20120330
'    sSql = sSql & " WHERE PP.cPersCod = '" & psPersCod & "' AND SUBSTRING(P.cCtaCod,6,3) IN (121, 221, 514)"
'        '& " Order By dEmitido"
'    sSql = sSql & " and ((PP.nPrdPersRelac=38 AND CCF.cConsorcio IS null) " 'WIOR 20120330
'    sSql = sSql & " or (PP.nPrdPersRelac=20 AND CCF.cConsorcio IS not null) " 'WIOR 20120330
'    sSql = sSql & " or (PP.nPrdPersRelac=20 AND CCF.cConsorcio IS null) or (PP.nPrdPersRelac not in (28,29,36) ))" 'WIOR 20120330

   sSql = "exec stp_sel_DatosPosicionClienteCF '" & psPersCod & "'" 'WIOR 20120804
    
    Conn.AbreConexion
    Set DatosPosicionClienteCF = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorDatosPosicionClienteCF:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Function DatosPosicionClienteComentarios(ByVal psPersCod As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String

    On Error GoTo ErrorDatosPosicionClienteComentario

    Set Conn = New COMConecta.DCOMConecta
    'CTI6-20210503-ERS032-2019. Comentó
'    sSql = "Select C.cPersCod, C.cMovNro, C.cComentario From Persona P JOIN PersComentario C ON " _
'        & "P.cPersCod = C.cPersCod WHERE P.cPersCod = '" & psPersCod & "' ORDER BY C.cMovNro DESC"
    sSQL = "exec stp_sel_DatosPosicionComentarios '" & psPersCod & "'"
    'Fin CTI6-20210503
    
    Conn.AbreConexion
    Set DatosPosicionClienteComentarios = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorDatosPosicionClienteComentario:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

'*******************************************************************
'Se Agrego para tener la Calificacion RCC en la Posicion del Cliente

Public Function DatosPosicionClienteCalificacionSBS(ByVal pbPersNat As Boolean, _
                                                    ByVal psPersDoc As String, _
                                                    ByRef psCodDeudorRCC As String, Optional psPersTDoc As String = "1") As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

'arlo20200312 begin
'    If pbPersNat Then
'        sSql = " SELECT R.Calif_0 as nNormal, R.Calif_1 as nPotencial, R.Calif_2 as nDeficiente, "
'        sSql = sSql & " R.Calif_3 as nDudoso, R.Calif_4 as nPerdido, Can_Ents,Fec_Rep, Cod_Edu  "
'        sSql = sSql & " FROM DBConsolidada..rccTotal R "
'        sSql = sSql & " WHERE (R.Cod_Doc_Id ='" & psPersDoc & "') "
'        sSql = sSql & "and (case R.Tip_Doc_Id when 1 then 1 "
'        sSql = sSql & "                      when 4 then 2 "
'        sSql = sSql & "                      when 5 then 11 end)=" & psPersTDoc
'    Else
'        sSql = " SELECT R.Calif_0 as nNormal, R.Calif_1 as nPotencial, R.Calif_2 as nDeficiente, "
'        sSql = sSql & " R.Calif_3 as nDudoso, R.Calif_4 as nPerdido, Can_Ents,Fec_Rep, Cod_Edu "
'        sSql = sSql & " FROM DBConsolidada..rccTotal R "
'        sSql = sSql & " WHERE (R.Cod_Doc_Trib ='" & psPersDoc & "') "
'    End If
'arlo20200312 end

    sSql = "exec stp_sel_DatosPosicionClienteCalificacionSBS '" & pbPersNat & "','" & psPersDoc & "'" ''arlo20200312"
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set rs = Conn.CargaRecordSet(sSql)
    If rs.EOF Then
        psCodDeudorRCC = ""
    Else
        psCodDeudorRCC = rs!cod_Edu
    End If
    Set DatosPosicionClienteCalificacionSBS = rs
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

End Function

Public Function DatosPosicionClienteEndeudamientoSBS(ByVal psCodDeudorRCC As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    
'arlo20200312 begin
'    sSql = " Select "
'    sSql = sSql & " Isnull(sum(Case when cod_Cuenta like '141[1456]%' then Val_Saldo end),0 ) DDirSoles, "
'    sSql = sSql & " Isnull(sum(Case when cod_Cuenta like '142[1456]%' then Val_Saldo end ),0) DDirDolar,"
'    sSql = sSql & " Isnull(sum(Case when cod_Cuenta like '7112%' then Val_Saldo end ),0) DIndSoles,"
'    sSql = sSql & " Isnull(sum(Case when cod_Cuenta like '7122%' then Val_Saldo end ),0) DIndDolar"
'    sSql = sSql & " From DBConsolidada..RccTotaldet"
'    sSql = sSql & " Where cod_Edu ='" & psCodDeudorRCC & "' "
'arlo20200312 end

    sSql = "exec stp_sel_DatosPosicionClienteEndeudamientoSBS '" & psCodDeudorRCC & "'" ''arlo20200312"
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set DatosPosicionClienteEndeudamientoSBS = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

End Function

Public Function DatosPosicionClienteCalificacionCMAC(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    
    'CTI6-20210503-ERS032-2019. Comentó
'    sSql = "SELECT TOP 3 CAT.dFecha, "
'    sSql = sSql & " ISNULL(CAT.cCalGen, '') as nCalFinal, ISNULL(CAT.cCalHist, '') as nCalRiesgos, ISNULL(CAT.cCalSistF, '') as nCalSistFinan "
'    sSql = sSql & " FROM DBConsolidada..CreditoAudiTotal CAT "
'    sSql = sSql & " WHERE CAT.ccodpers='" & psPersCod & "' "
'    sSql = sSql & " ORDER BY CAT.ccodpers, convert(varchar(8), CAT.dFecha,112) desc"
    sSQL = "exec stp_sel_DatosPosicionClienteCalificacionCMAC '" & psPersCod & "'"
    'Fin CTI6-20210503
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set DatosPosicionClienteCalificacionCMAC = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

End Function
'********************************************************************

Public Sub NuevoPermisoNivelApr(ByVal cCodNiv As String, ByVal cPersCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorNuevoPermisoNivelApr
    
    sSql = "INSERT INTO ColocCredPersNivelesApr(cCodNiv,cPersCod)"
    sSql = sSql & " VALUES('" & cCodNiv & "','" & cPersCod & "')"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoPermisoNivelApr:
    Err.Raise Err.Number, "Nuevo Permiso de Nivel de Aprobacion", Err.Description
End Sub

Public Sub EliminarPermisoNivelApr(ByVal cCodNiv As String, ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorEliminarPermisoNivelApr
    sSql = "DELETE ColocCredPersNivelesApr WHERE cCodNiv = '" & cCodNiv & "' AND cPersCod = '" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorEliminarPermisoNivelApr:
End Sub

Public Sub NuevoNivelAprobacion(ByVal cCodNiv As String, ByVal cProduct As String, _
    ByVal cTpoCred As String, ByVal cNivel As String, ByVal nMontoMax As Double, _
    ByVal nMontoMin As Double, ByVal pnMoneda As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorNuevoNivelAprobacion
    
    sSql = "INSERT INTO ColocCredNivelesApr(cCodNiv, cProduct, cTpoCred, nNivel, nMontoMax, nMontoMin, nMoneda)"
    sSql = sSql & " VALUES('" & cCodNiv & "','" & cProduct & "','" & cTpoCred & "'," & Trim(cNivel) & "," & Format(nMontoMax, "#0.00") & "," & Format(nMontoMin, "#0.00") & "," & pnMoneda & ")"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoNivelAprobacion:
    Err.Raise Err.Number, "Nuevo Nivel de Aprobacion", Err.Description
End Sub


Public Sub EliminaNivelAprobacion(ByVal cCodNiv As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaNivelAprobacion
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    sSql = "DELETE ColocCredPersNivelesApr Where cCodNiv = '" & cCodNiv & "'"
    oConecta.ConexionActiva.Execute sSql
    sSql = "DELETE ColocCredNivelesApr WHERE cCodNiv = '" & cCodNiv & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorEliminaNivelAprobacion:
    Err.Raise Err.Number, "Eliminar un Nivel de Aprobacion", Err.Description
    
End Sub

Public Function RecuperaCreditosArchivoPagoLote(ByVal MatDatos As Variant, ByVal psCodInst As String, _
                                            Optional ByVal pbIncluyeMora As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadCodMod As String
Dim i As Integer

    On Error GoTo ErrorRecuperaCreditosPagoLote
    sCadCodMod = ""
    For i = 0 To UBound(MatDatos) - 1
        sCadCodMod = sCadCodMod & "'" & MatDatos(i) & "',"
    Next i
    sCadCodMod = Mid(sCadCodMod, 1, Len(sCadCodMod) - 1)
    sSql = "Select '',P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, SUM(CD.nMonto) nPago, "
    sSql = sSql & "Pers.cPersNombre, SUM(CD.nMonto - CD.nMontoPagado) nSaldoCuota, "
    sSql = sSql & " nMora = (Select SUM(CD2.nMonto - CD2.nMontoPagado) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
    sSql = sSql & "CC.nDiasAtraso, "
    'sSql = sSql & " SUM(CD.nMonto - CD.nMontoPagado) nDeudaTotal, "
    sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen), "
    sSql = sSql & " CC.cMetLiquidacion, "
    sSql = sSql & "P.nTransacc, P.nPrdEstado "
    sSql = sSql & " ,Vencim  = (Select CONVERT(varchar(10),CCal.dVenc,103)  From ColocCalendario CCal WHERE CCal.cCtaCod = P.cCtaCod AND CCal.nColocCalendApl = 1 AND CCal.nNroCalen = CC.nNroCalen  AND CCal.nCuota = CC.nNroProxCuota ) "
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSql = sSql & "             And CV.cCodModular in (" & sCadCodMod & ") "
    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosArchivoPagoLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosPagoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCreditosCodigoModular(ByVal psCodInst As String, _
    ByVal psAgencia As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaCreditosCodigoModular
    sSql = "Select ' ',P.cCtacod, CV.cCodModular, "
    sSql = sSql & "Pers.cPersNombre, ISNULL(CV.cCargo,'') as cCargo, ISNULL(CV.cCARBEN,'') as cCARBEN,ISNULL(CV.cT_Plani,'') as cT_Plani  "
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    'sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' "
    sSql = sSql & " AND SUBSTRING(P.cCtaCod,4,2) in " & psAgencia
    sSql = sSql & " Order by Pers.cPersNombre "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCodigoModular = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosCodigoModular:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCreditosPagoLote(ByVal psCodInst As String, Optional ByVal pnMoneda As Moneda = gMonedaNacional, _
                                           Optional ByVal pbIncluyeMora As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaCreditosPagoLote
    'Comentado Por MAVM 20120209 ***
    '  sSql = " SELECT * FROM ( "
    '    sSql = sSql & " Select ' ' as cOK,P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, "
    '    sSql = sSql & " ROUND(SUM(CD.nMonto-CD.nMontoPagado),6) nPago,"
    '    sSql = sSql & "Pers.cPersNombre, ROUND(SUM(CD.nMonto - CD.nMontoPagado),2) nSaldoCuota, "
    '    sSql = sSql & " nMora = (Select ROUND(SUM(CD2.nMonto - CD2.nMontoPagado),2) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
    '    sSql = sSql & " CC.nDiasAtraso, "
    '
    '    'ARCV 30-04-2006
    '
    '    'ARCV 06-03-2007
    '    'sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen ), "
    '    'sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen AND nPrdConceptoCod NOT LIKE '12%' AND nPrdConceptoCod <>1101 )+ dbo.ColocCred_ObtieneGastoFechaCredito_PL(P.cCtaCod), "
    '    '---------
    '    If pbIncluyeMora Then
    '        sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen AND nPrdConceptoCod NOT LIKE '12%')+ dbo.ColocCred_ObtieneGastoFechaCredito_PL(P.cCtaCod), "
    '    Else
    '        sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen AND nPrdConceptoCod NOT LIKE '12%' AND nPrdConceptoCod <>1101 )+ dbo.ColocCred_ObtieneGastoFechaCredito_PL(P.cCtaCod), "
    '    End If
    '    '--------------
    '
    '    sSql = sSql & " CC.cMetLiquidacion, "
    '    sSql = sSql & "P.nTransacc, P.nPrdEstado, CC.nNroCalen, "
    '    sSql = sSql & "dbo.ITFImpuestoRedondeo(ROUND( dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6)) as nITF,   "
    '
    '    '**Modificado por DAOR 20080319 ****************************************************************************************************************************
    '    'sSql = sSql & " nCapital = (Select ROUND(SUM(CD3.nMonto - CD3.nMontoPagado),2) From ColocCalendDet CD3 WHERE CD3.cCtaCod = P.cCtaCod AND CD3.nColocCalendApl = 1 AND CD3.nNroCalen = CC.nNroCalen  AND CD3.nPrdConCeptoCod = 1000 AND CD3.nCuota = CC.nNroProxCuota), "
    '    'sSql = sSql & " nInteres = (Select ROUND(SUM(CD4.nMonto - CD4.nMontoPagado),2) From ColocCalendDet CD4 WHERE CD4.cCtaCod = P.cCtaCod AND CD4.nColocCalendApl = 1 AND CD4.nNroCalen = CC.nNroCalen  AND CD4.nPrdConCeptoCod = 1100 AND CD4.nCuota = CC.nNroProxCuota), "
    '    'sSql = sSql & " Protesto = (Select ISNULL(ROUND(SUM(CD6.nMonto - CD6.nMontoPagado),2),0) From ColocCalendDet CD6 WHERE CD6.cCtaCod = P.cCtaCod AND CD6.nColocCalendApl = 1 AND CD6.nNroCalen = CC.nNroCalen  AND CD6.nPrdConCeptoCod IN (1219,1220) and CD6.nCuota = CC.nNroProxCuota ), "
    '    ''ARCV 22-06-2007
    '    ''sSql = sSql & " Gastos   = (Select ISNULL(ROUND(SUM(CD5.nMonto - CD5.nMontoPagado),2),0) From ColocCalendDet CD5 WHERE CD5.cCtaCod = P.cCtaCod AND CD5.nColocCalendApl = 1 AND CD5.nNroCalen = CC.nNroCalen  AND CD5.nPrdConCeptoCod not in (1000,1100,1101,1219,1220) AND CD5.nCuota = CC.nNroProxCuota ) "
    '    'sSql = sSql & " Gastos   = (Select ISNULL(ROUND(SUM(CD5.nMonto - CD5.nMontoPagado),2),0) From ColocCalendDet CD5 WHERE CD5.cCtaCod = P.cCtaCod AND CD5.nColocCalendApl = 1 AND CD5.nNroCalen = CC.nNroCalen  AND CD5.nPrdConCeptoCod LIKE '12%' AND CD5.nCuota = CC.nNroProxCuota ) "
    '    ''-------
    '    sSql = sSql & " nCapital = (round(sum(case when CD.nPrdConCeptoCod = 1000 then CD.nMonto - CD.nMontoPagado else 0 end),2)), "
    '    sSql = sSql & " nInteres = (round(sum(case when CD.nPrdConCeptoCod = 1100 then CD.nMonto - CD.nMontoPagado else 0 end),2)), "
    '    sSql = sSql & " Protesto = (isnull(round(sum(case when CD.nPrdConCeptoCod IN (1219,1220) then CD.nMonto - CD.nMontoPagado else 0 end),2),0)), "
    '    sSql = sSql & " Gastos   = (isnull(round(sum(case when CD.nPrdConCeptoCod LIKE '12%' then CD.nMonto - CD.nMontoPagado else 0 end),2),0)) "
    '    '***********************************************************************************************************************************************************
    '
    '    'ARCV 30-10-2006
    '    'ARCV 22-06-2007
    '    'sSql = sSql & " ,Vencim  = (Select CONVERT(varchar(10),CCal.dVenc,103)  From ColocCalendario CCal WHERE CCal.cCtaCod = P.cCtaCod AND CCal.nColocCalendApl = 1 AND CCal.nNroCalen = CC.nNroCalen  AND CCal.nCuota = CC.nNroProxCuota ) "
    '    sSql = sSql & " ,Vencim  = CONVERT(VarChar(10), CCal.dVenc, 103) "
    '    sSql = sSql & " ,CE.nPlazo, "
    '    sSql = sSql & "ROUND( dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) as nITFCalc "
    '    '---------------
    '    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    '    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    '    'ARCV 22-06-2007
    '    sSql = sSql & " INNER JOIN ColocCalendario CCAL ON CCAL.cCtaCod = P.cCtaCod AND CCAL.nColocCalendApl = 1 AND CCAL.nNroCalen = CC.nNroCalen AND CCAL.nCuota = CC.nNroProxCuota "
    '    '---------
    '    'ARCV 30-04-2007
    '    'sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    '    If pbIncluyeMora Then
    '        sSql = sSql & "             Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota "
    '    Else
    '        sSql = sSql & "             Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    '    End If
    '    '----------------
    '    'ARCV 22-06-2007
    '    sSql = sSql & "             INNER JOIN ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = 2002 "
    '    '------
    '    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    '    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    '    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "'  And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") and SUBSTRING(P.CCTACOD,9,1)='" & pnMoneda & "' "
    '    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen ,CCal.dVenc,CE.nPlazo " 'ARCV 22-06-2007 ,CCal.dVenc,CE.nPlazo
    '    sSql = sSql & " HAVING SUM(CD.nMonto - CD.nMontoPagado) > 0  ) AS PAGOSLOTE "
    '    sSql = sSql & " ORDER BY cPersNombre "
    '
    ''    sSql = "Select ' ',P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, "
    ''    sSql = sSql & " ROUND(SUM(CD.nMonto-CD.nMontoPagado) + dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) nPago,"
    ''    sSql = sSql & "Pers.cPersNombre, ROUND(SUM(CD.nMonto - CD.nMontoPagado),2) nSaldoCuota, "
    ''    sSql = sSql & " nMora = (Select ROUND(SUM(CD2.nMonto - CD2.nMontoPagado),2) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
    ''    sSql = sSql & " CC.nDiasAtraso, "
    ''    sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen ), "
    ''    sSql = sSql & " CC.cMetLiquidacion, "
    ''    sSql = sSql & "P.nTransacc, P.nPrdEstado, CC.nNroCalen, "
    ''    sSql = sSql & "ROUND(dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) as nITF,   "
    ''    sSql = sSql & " nCapital = (Select ROUND(SUM(CD3.nMonto - CD3.nMontoPagado),2) From ColocCalendDet CD3 WHERE CD3.cCtaCod = P.cCtaCod AND CD3.nColocCalendApl = 1 AND CD3.nNroCalen = CC.nNroCalen  AND CD3.nPrdConCeptoCod = 1000 AND CD3.nCuota = CC.nNroProxCuota), "
    ''    sSql = sSql & " nInteres = (Select ROUND(SUM(CD4.nMonto - CD4.nMontoPagado),2) From ColocCalendDet CD4 WHERE CD4.cCtaCod = P.cCtaCod AND CD4.nColocCalendApl = 1 AND CD4.nNroCalen = CC.nNroCalen  AND CD4.nPrdConCeptoCod = 1100 AND CD4.nCuota = CC.nNroProxCuota), "
    ''    sSql = sSql & " Protesto = (Select ISNULL(ROUND(SUM(CD6.nMonto - CD6.nMontoPagado),2),0) From ColocCalendDet CD6 WHERE CD6.cCtaCod = P.cCtaCod AND CD6.nColocCalendApl = 1 AND CD6.nNroCalen = CC.nNroCalen  AND CD6.nPrdConCeptoCod IN (1219,1220) and CD6.nCuota = CC.nNroProxCuota ), "
    ''    sSql = sSql & " Gastos   = (Select ISNULL(ROUND(SUM(CD5.nMonto - CD5.nMontoPagado),2),0) From ColocCalendDet CD5 WHERE CD5.cCtaCod = P.cCtaCod AND CD5.nColocCalendApl = 1 AND CD5.nNroCalen = CC.nNroCalen  AND CD5.nPrdConCeptoCod not in (1000,1100,1101,1219,1220) AND CD5.nCuota = CC.nNroProxCuota ) "
    ''    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    ''    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    ''    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    ''    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    ''    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    ''    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "'  And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") and SUBSTRING(P.CCTACOD,9,1)='" & pnMoneda & "' "
    ''    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
    ''    sSql = sSql & " HAVING SUM(CD.nMonto - CD.nMontoPagado) > 0  "
    ''    sSql = sSql & " ORDER BY  Pers.cPersNombre"
    sSql = " exec stp_sel_RecuperaCreditosPagoLote '" & psCodInst & "','" & pnMoneda & "'," & IIf(pbIncluyeMora = True, 1, 0)
    '***

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPagoLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosPagoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosExtornoGeneral(ByVal pdFecOper As Date, ByVal pTipoExt As TipoExtorno, Optional psCodusu As String = "", Optional psCtaCod As String = "", Optional psCodCli As String = "", _
    Optional ByVal pbDesemb As Boolean = False, Optional ByVal psCodAge As String = "", Optional ByVal psCodOpera As String = "", Optional ByVal pbPagoHonra As Boolean = False, Optional ByVal pbExtornoVigen As Integer = 0) As ADODB.Recordset
    'WIOR 20131228 AGREGO pbPagoHonra
    'LUCV 20160714 AGREGO pbExtornoVigen
Dim sSql As String
Dim sFecOperacion As String 'LUCV20160615, ERS004-2016
sFecOperacion = Format(pdFecOper, "yyyymmdd") 'LUCV20160615, ERS004-2016

Dim oConecta As COMConecta.DCOMConecta

'ARCV 04-09-2006
'Modificado para manejar los Extornos de Desembolsos con Abono a Cuenta
''MADM 20100915 -- modificacion para evitar extornos individuales a pago en lote
    On Error GoTo ErrorRecuperaDatosExtornoGeneral
    'LUCV20160615, Comentó segun ERS2016-004 **********
'    Select Case pTipoExt
'        Case gTipoExtornoCuenta
'            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
'            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
'            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
'            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And MC.cCtaCod = '" & psCtaCod & "'"
'            sSql = sSql & "       And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN Mov M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' " 'AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
'            sSql = sSql & "         AND (M1.NMOVFLAG<>'2' OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M1.nMovNro AND MC1.cOpeCod IN('100102','100103') )>0) "
'            sSql = sSql & "     Group by cCtaCod) "
'            sSql = sSql & " AND (M.nMovFlag <> 2 OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M.nMovNro AND MC.cOpeCod IN('100102','100103'))>0 ) "
'            'sSql = sSql & "       AND MC.cOpeCod <> '100301' "
'            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
'            If Not pbPagoHonra Then 'WIOR 20131228
'                If pbDesemb Then
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
'                Else
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
'                End If
'             'WIOR 20131228 ********************************
'            Else
'                sSql = sSql & "       AND MC.cOpeCod = '" & gCredPagHonramiento & "' "
'            End If
'            'WIOR FIN *************************************
'            'ALPA 201603224
'            sSql = sSql & " AND NOT (O.cOpeCod LIKE '9901%' OR O.cOpeCod LIKE '9903%' OR O.cOpeCod LIKE '107[1234]%' or O.cOpeCod LIKE '10081[0123456]%')"
'            sSql = sSql & " AND NOT (M.cOpeCod =  '102100')" 'MADM 20100915
'            sSql = sSql & " AND NOT (MC.cOpeCod LIKE '10092[01]%') " 'EJVG20140326
'            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
'            sSql = sSql & " Order By M.nMovNro DESC "
'
'        Case gTipoExtornoUsuario
'            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago"
'            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
'            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
'            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And Right(M.cMovNro,4) = '" & psCodusu & "'"
'            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOVDIARIO M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' " 'AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
'            sSql = sSql & "         AND (M1.NMOVFLAG<>'2' OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M1.nMovNro AND MC1.cOpeCod IN('100102','100103') )>0) "
'            sSql = sSql & "         Group by cCtaCod) "
'            sSql = sSql & " AND (M.nMovFlag <> 2 OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M.nMovNro AND MC.cOpeCod IN('100102','100103'))>0 ) "
'            'sSql = sSql & "         AND MC.cOpeCod <> '100301' " 'Comentado para poder extornar los pagos de credito morosos gitu 2009-07-01
'            'sSQL = sSQL & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
'            If Not pbPagoHonra Then 'WIOR 20131228
'                If pbDesemb Then
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
'                Else
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
'                End If
'             'WIOR 20131228 ********************************
'            Else
'                sSql = sSql & "       AND MC.cOpeCod = '" & gCredPagHonramiento & "' "
'            End If
'            'WIOR FIN *************************************
'            'ALPA 201603224
'            sSql = sSql & " AND NOT (O.cOpeCod LIKE '9901%' OR O.cOpeCod LIKE '9903%' OR O.cOpeCod LIKE '107[1234]%' or O.cOpeCod LIKE '10081[0123456]%')"
'            sSql = sSql & " AND NOT (M.cOpeCod =  '102100')" 'MADM 20100915
'            sSql = sSql & " AND NOT (MC.cOpeCod LIKE '10092[01]%') " 'EJVG20140326
'            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
'            sSql = sSql & " Order By M.nMovNro DESC "
'
'        Case gTipoExtornoCliente
'            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
'            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
'            sSql = sSql & "     Inner Join ProductoPersona PP ON PP.cCtaCod =  MC.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
'            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
'            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And PP.cPersCod = '" & psCodCli & "'"
'            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOVDIARIO M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' " 'AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
'            sSql = sSql & "         AND (M1.NMOVFLAG<>'2' OR (SELECT COUNT(*)FROM MovRef WHERE nMovNroRef= M1.nMovNro AND MC1.cOpeCod IN('100102','100103'))>0) "
'            sSql = sSql & "         Group by cCtaCod) "
'            sSql = sSql & " AND (M.nMovFlag <> 2 OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M.nMovNro AND MC.cOpeCod IN('100102','100103'))>0 ) "
'            sSql = sSql & "         AND MC.cOpeCod <> '100301' "
'            'sSQL = sSQL & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
'            If Not pbPagoHonra Then 'WIOR 20131228
'                If pbDesemb Then
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
'                Else
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
'                End If
'             'WIOR 20131228 ********************************
'            Else
'                sSql = sSql & "       AND MC.cOpeCod = '" & gCredPagHonramiento & "' "
'            End If
'            'WIOR FIN *************************************
'            'ALPA 201603224
'            sSql = sSql & " AND NOT (O.cOpeCod LIKE '9901%' OR O.cOpeCod LIKE '9903%' OR O.cOpeCod LIKE '107[1234]%' or O.cOpeCod LIKE '10081[0123456]%')"
'            sSql = sSql & " AND NOT (M.cOpeCod =  '102100')" 'MADM 20100915
'            sSql = sSql & " AND NOT (MC.cOpeCod LIKE '10092[01]%') " 'EJVG20140326
'            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
'            sSql = sSql & " Order By M.nMovNro DESC "
'
'        Case gTipoExtornoGeneral
'            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
'            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
'            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
'            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' "
'            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOV M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' " 'AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
'            sSql = sSql & "         AND (M1.NMOVFLAG<>'2' OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M1.nMovNro AND MC1.cOpeCod IN('100102','100103') )>0) "
'            sSql = sSql & "         Group by cCtaCod) "
'            sSql = sSql & " AND (M.nMovFlag <> 2 OR (SELECT COUNT(*)FROM MovRef MR INNER JOIN Mov ON MR.nMovNro=Mov.nMovNro AND Mov.nMovFlag<>2 WHERE MR.nMovNroRef= M.nMovNro AND MC.cOpeCod IN('100102','100103'))>0 ) "
'            sSql = sSql & "         AND MC.cOpeCod <> '100301' "
'            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
'            If Not pbPagoHonra Then 'WIOR 20131228
'                If pbDesemb Then
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
'                Else
'                    sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
'                End If
'             'WIOR 20131228 ********************************
'            Else
'                sSql = sSql & "       AND MC.cOpeCod = '" & gCredPagHonramiento & "' "
'            End If
'            'WIOR FIN *************************************
'            'ALPA 201603224
'            sSql = sSql & " AND NOT (O.cOpeCod LIKE '9901%' OR O.cOpeCod LIKE '9903%' OR O.cOpeCod LIKE '107[1234]%' or O.cOpeCod LIKE '10081[0123456]%')"
'            sSql = sSql & " AND NOT (M.cOpeCod =  '102100')" 'MADM 20100915
'            sSql = sSql & " AND NOT (MC.cOpeCod LIKE '10092[01]%') " 'EJVG20140326
'            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
'            sSql = sSql & " Order By M.nMovNro DESC "
'
'    End Select
    
'Fin Comentado-> LUCV20160615 **********
sSql = " exec stp_sel_RecuperaDatosExtornoGeneral '" & sFecOperacion & "', '" & pTipoExt & "',  '" & psCodusu & "', '" & psCtaCod & "', '" & psCodCli & "', " & IIf(pbDesemb = True, 1, 0) & ", '" & psCodAge & "', " & IIf(pbPagoHonra = True, 1, 0) & ", " & pbExtornoVigen & ", '" & gCredPagHonramiento & "', '" & gColRelPersTitular & "'"
    
    

    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    'oConecta.ConexionActiva.CommandTimeout = 60000
    oConecta.ConexionActiva.CommandTimeout = 500 ' ARCV 06-09-2006
    
    Set RecuperaDatosExtornoGeneral = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosExtornoGeneral:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaPagosEnLote(ByVal pdFecOper As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "Select M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) nMonto, RIGHT(M.cMovNro,4) cUsuario, cOpeCod='', substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora "
    sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
    sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' "
    sSql = sSql & "         And M.cOpeCod =  '" & gCredPagLote & "'"
    ' " And MC.nMovNro in (Select MAX(MC2.nMovNro) From MovCol MC2 Inner join Mov M2 On  MC2.nMovNro = M2.nMovNro Where MC2.cCtaCod = MC.cCtaCod and MC2.copeCod not in ('" & gCredExtPago & "','" & gCredExtPagoLote & "')  And M2.nMovFlag = 0  Group by MC2.cCtaCod) "
    sSql = sSql & "         And M.nMovFlag = " & gMovFlagVigente
    sSql = sSql & "         And MC.cOpeCod LIKE '100[234567]%'" 'EJVG20140812
    sSql = sSql & " group by M.cMovDesc,MC.nMovNro,RIGHT(M.cMovNro,4), substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2)"
    sSql = sSql & " Order By MC.nMovNro DESC "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagosEnLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosdePagoEnLote(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    'sSql = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro & " and NOT cOpecod LIKE '99%'"
    sSql = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro & " and cOpecod LIKE '100[234567]%'" 'EJVG20140813
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosdePagoEnLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function SaldoConsolCapta(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    'sSql = "SELECT SUM(nSaldo) as SaldoCaptac FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
    '    & "WHERE SUBSTRING(P.cCtaCod,6,3) IN (232,233,234,239) " _
    '    & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")" _
    '    & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 10 "

    'EJVG20120430
    sSql = "Exec stp_sel_SaldoConsolCapta '" & psPersCod & "'," & pnMoneda
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolCapta = IIf(IsNull(rs!SaldoCaptac), 0, rs!SaldoCaptac)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function SaldoConsolColoc(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    'sSql = "SELECT SUM(nSaldo) SaldoColoc FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
    '    & "WHERE SUBSTRING(P.cCtaCod,6,3) NOT IN (232,233,234,239,121,221) " _
    '    & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado NOT IN (" & gColocEstSolic & "," & gColocEstSug _
    '    & "," & gColocEstAprob & "," & gColocEstCancelado & "," & gColocEstRech & "," & gColocEstRefinanc & "," & gColocEstRetirado _
    '    & "," & gColPEstRegis & "," & gColPEstDifer & "," & gColPEstCance & "," & gColPEstRemat & "," & gColPEstAdjud _
    '    & "," & gColPEstSubas & "," & gColPEstAnula & "," & gColPEstChafa & "," & gColPEstAnNoD & "," & gPigEstRegis _
    '    & "," & gPigEstCancelPendRes & "," & gPigEstRescate & "," & gPigEstRematPFact & "," & gPigEstRematFact & " " _
    '    & "," & gPigEstAdjud & "," & gPigEstAnula & ")" _
    '    & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 20 "
    
    'EJVG20120430
    sSql = "Exec stp_sel_SaldoConsolColoc '" & psPersCod & "'," & pnMoneda
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolColoc = IIf(IsNull(rs!SaldoColoc), 0, rs!SaldoColoc)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'EJVG20120430
Public Function SaldoConsolColocCFGarantiza(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_SaldoConsolColocCFGarantiza '" & psPersCod & "'," & pnMoneda
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolColocCFGarantiza = IIf(IsNull(rs!SaldoCF), 0, rs!SaldoCF)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'END EJVG
Public Function SaldoConsolCF(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    'sSql = "SELECT SUM(nSaldo) SaldoCF FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
    '    & "WHERE SUBSTRING(P.cCtaCod,6,3) IN (121,221)" _
    '    & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado IN (" & gColocEstVigNorm & ")" _
    '    & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 20"
    
    'EJVG20120430
    sSql = "Exec stp_sel_SaldoConsolCF '" & psPersCod & "'," & pnMoneda
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolCF = IIf(IsNull(rs!SaldoCF), 0, rs!SaldoCF)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Sub AgregaPersComentario(ByVal sPersCod As String, ByVal sMovNro As String, _
        ByVal sComentario As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizaNivel
    sSql = "INSERT PersComentario (cPersCod,cMovNro,cComentario) " _
        & "Values ('" & sPersCod & "','" & sMovNro & "','" & sComentario & "')"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizaNivel:
    Err.Raise Err.Number, "Agrega Persona Comentario", Err.Description
End Sub


'Carga los Datos necesarios para los Niveles de Aprobacion de Creditos

Public Function Cargar_Datos_Generales(ByRef psApoderados As String, ByRef prsAnalista As ADODB.Recordset, _
                                       ByRef prsNivAprob As ADODB.Recordset, ByRef prsNivel As ADODB.Recordset, _
                                       ByRef prsNivAdm As ADODB.Recordset, ByRef prsProducto As ADODB.Recordset, _
                                       ByRef prsTipoCred As ADODB.Recordset, ByRef prsMoneda As ADODB.Recordset)

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oCons As COMDConstantes.DCOMConstantes
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    
    psApoderados = oGen.LeeConstSistema(gConstSistRHCargoCodApoderados)
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " AND nRHEstado = 201 "
    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    sSql = sSql & " where  RC.cRHCargoCod in (" & psApoderados & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod)"
    sSql = sSql & " order by P.cPersNombre "
    
    Set prsAnalista = oConecta.CargaRecordSet(sSql)
    
    Set prsNivAprob = RecuperaNivelesAprobacion
    
    Set prsNivel = RecuperaNiveles
    
    sSql = "Select nNivel,cDescripcion from ColocCredNivelesTipos"
    Set prsNivAdm = oConecta.CargaRecordSet(sSql)
    
    sSql = "select nConsValor, cConsDescripcion from Constante where convert(varchar(15),nConsValor) not like '23_' AND nConsValor <> 305 AND nConsValor <> nConsCod AND nConsCod = " & gProducto
    Set prsProducto = oConecta.CargaRecordSet(sSql)
    
    'Cargar Tipo de Credito
    Set oCons = New COMDConstantes.DCOMConstantes
    Set prsTipoCred = oCons.RecuperaConstantes(gColocCredTipo)
       
    'Carga Moneda
    Set prsMoneda = oCons.RecuperaConstantes(gMoneda)
    
    oConecta.CierraConexion
    Set oCons = Nothing
    Set oGen = Nothing
    Set oConecta = Nothing
    Exit Function
ErrorCargar_Datos_Generales:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description

End Function

'MADM 20100128  --------------------------------------------------
'madm 20100201-----------------------------------------------------------------------
Public Function CargaDatos_Comite_ult(ByVal id_agencia As Integer, ByVal nom_comite As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORCargaDatos_Comite_ult
    
    sSql = " exec stp_sel_listaComitexAgeyNom " & id_agencia & ", '" & nom_comite & "'"
           
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaDatos_Comite_ult = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaDatos_Comite_ult:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function CargaDatos_Comite(ByVal id_comite As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORCargaDatos_Comite
    
    sSql = " exec stp_sel_listaComitexCod " & id_comite & ""
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaDatos_Comite = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaDatos_Comite:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function Num_Comite_Agencia(ByVal codage As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORNum_Comite_Agencia
    
    sSql = " exec stp_sel_listaNumComitexCodAge " & codage & ""
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Num_Comite_Agencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORNum_Comite_Agencia:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function CargaAnalistasAgencia_Det(ByVal id_codage As Integer, id_comite As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORCargaAnalistasDet
    
    sSql = " exec stp_sel_listaCodPerAnalistaComite " & id_codage & "," & id_comite & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaAnalistasAgencia_Det = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaAnalistasDet:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'no aprece el coordinador
Public Function CargaAnalistasAgencia(ByVal codage As String, Optional filtro As Boolean = False) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim sSqlaux As String
Dim oGen As COMDConstSistema.DCOMGeneral
    

    On Error GoTo ERRORCargaAnalistas
    'FRHU 20150326
    'Set oGen = New COMDConstSistema.DCOMGeneral
    'sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    'Set oGen = Nothing
    
    'If filtro = True Then
        'sSqlaux = sSql & " AND R.cPersCod not in (select id_analista cPersCod from Comite_det)"
    'End If
    'sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    'sSql = sSql & " AND nRHEstado = 201 "
    'sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    'sSql = sSql & " where  RC.cRHCargoCod in(" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    'sSql = sSql & " AND R.cPersCod not in (select res_comite from comite) "
    'sSql = sSql & " AND R.cAgenciaAsig = " & codage & " " & sSqlaux
    'sSql = sSql & " order by P.cPersNombre "
    sSql = "stp_sel_CargaAnalistaAgencia '" & Trim(codage) & "'," & IIf(filtro = True, 1, 0)
    'FIN FRHU 20150326
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set CargaAnalistasAgencia = R
    Set R = Nothing
    Exit Function
    
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'no responsable ni analistas de otros comites
Public Function CargaAnalistasAgenciasinComite(ByVal codage As String, ByVal codmite As String, Optional filtro As Boolean = False) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim sSqlaux As String
Dim oGen As COMDConstSistema.DCOMGeneral

    On Error GoTo ERRORCargaAnalistasAgenciasinComite
    'FRHU 20150326
    'Set oGen = New COMDConstSistema.DCOMGeneral
    'sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    'Set oGen = Nothing
       
    'sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    'sSql = sSql & " AND nRHEstado = 201 "
    'sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    'sSql = sSql & " where  RC.cRHCargoCod in(" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    'sSql = sSql & " AND R.cPersCod not in (select res_comite from comite) "
    'sSql = sSql & " AND R.cPersCod not in (select id_analista cPersCod from Comite_det where id_comite <> " & codmite & " )"
    'sSql = sSql & " AND R.cAgenciaAsig = " & codage & " "
    'sSql = sSql & " order by P.cPersNombre "
    sSql = "stp_sel_CargaAnalistasAgenciasinComite '" & codage & "'," & CInt(codmite)
    'FIN FRHU 20150326
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set CargaAnalistasAgenciasinComite = R
    Set R = Nothing
    Exit Function
    
ERRORCargaAnalistasAgenciasinComite:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

'madm reporte en excel ------------------------------------------------------
Public Function CargaPersonasComite(ByVal id_comite As Integer) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    sSql = " exec stp_sel_listaAnalistaxComite " & id_comite & ""
 
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaPersonasComite = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
   
    Exit Function
End Function
'end madm

Public Function CargaPersonasJefeyResponsable(ByVal id_comite As Integer, cod_age As Integer) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sSqlaux As String

    On Error GoTo ERRORCargaPeronasComite
     
     sSql = " exec stp_sel_listaCoordinadoryJefeagencia " & id_comite & ", " & cod_age & ""
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaPersonasJefeyResponsable = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaPeronasComite:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

'------------------------------------------------------------------------
Public Function CargaComiteAgencia(ByVal id_agencia As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORCargaComiteAgencia
    
   sSql = " exec stp_sel_CargaComitexAgencia " & id_agencia & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaComiteAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaComiteAgencia:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function CargaComiteAgenciaxComite(ByVal id_agencia As Integer, id_comite As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ERRORCargaComiteAgencia
    
     sSql = " exec stp_sel_CargaComitexAgenciaxComite " & id_agencia & ", " & id_comite & ""
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaComiteAgenciaxComite = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaComiteAgencia:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'end madm -20100202------------------------------------------------------------------------------------------

Public Function EliminacionComite(ByVal pIdComite As Integer) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    On Error GoTo ErrHandler
        sSql = "Delete From comite Where id_comite=" & pIdComite & ""
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        oConec.ConexionActiva.Execute sSql
        oConec.CierraConexion
        Set oConec = Nothing
        EliminacionComite = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    EliminacionComite = False
End Function
Public Function EliminacionComitedet(ByVal pIdComite As Integer) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    On Error GoTo ErrHandler
        
        sSql = "Delete From comite_det Where id_comite=" & pIdComite & ""
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        oConec.ConexionActiva.Execute sSql
        oConec.CierraConexion
        Set oConec = Nothing
        EliminacionComitedet = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    EliminacionComitedet = False
End Function

Public Sub Insertar_Datos_Comite(ByVal id_agencia As Integer, ByVal nom_comite As String, res_comite As String, fec_comite_det As Date)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Comite
   
   sSql = "Exec stp_ins_comite_nuevo " & id_agencia & ",'" & nom_comite & "','" & res_comite & "','" & Format(fec_comite_det, "mm/dd/yyyy") & "'"
          
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertar_Datos_Comite:
    Err.Raise Err.Number, "Insertar Datos Comite", Err.Description
End Sub

Public Sub Actualizar_Datos_Comite(id_comite As Integer, ByVal id_agencia As Integer, ByVal nom_comite As String, res_comite As String, fec_comite_det As Date)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Comite
   
   sSql = "Exec stp_upd_comite " & id_comite & "," & id_agencia & ",'" & nom_comite & "','" & res_comite & "','" & Format(fec_comite_det, "mm/dd/yyyy") & "'"
          
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertar_Datos_Comite:
    Err.Raise Err.Number, "Insertar Datos Comite", Err.Description
End Sub

Public Sub Insertar_Datos_Comite_det(id_comite As Integer, ByVal id_agencia As Integer, ByVal id_analista As String, ByVal nContador As Integer)
     Dim sSql As String
     Dim oConecta As COMConecta.DCOMConecta
     On Error GoTo ErrorInsertar_Datos_Comite_det
    
     sSql = "Exec stp_ins_comite_det_nuevo " & id_comite & ", " & id_agencia & ",'" & id_analista & "'," & nContador
        
     Set oConecta = New COMConecta.DCOMConecta
     oConecta.AbreConexion
     oConecta.ConexionActiva.Execute sSql
     oConecta.CierraConexion
     Set oConecta = Nothing
     Exit Sub
ErrorInsertar_Datos_Comite_det:
    Err.Raise Err.Number, "Insertar Datos Comite", Err.Description
End Sub
'end madm 20100202 -*-----------------------------------------------------

'Cargar otras Relaciones de Personas Con un Credito
Public Function ObtenerOtrasRelaciones(ByVal psPersCod As String) As ADODB.Recordset

Dim sSql As String
Dim sCtaCod As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorObtenerOtrasRelaciones

    'Obtenemos Otras Relaciones de Creditos Anteriores
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        
    sSql = "Select top 1 C.dVigencia, PP.* "
    sSql = sSql & " from ProductoPersona PP"
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = PP.cCtaCod"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = PP.cCtaCod"
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "' AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    sSql = sSql & " Order by C.dVigencia DESC"
        
    Set R = oConecta.CargaRecordSet(sSql)
            
    If Not R.EOF Then
        sCtaCod = R!cCtaCod
    Else
        sCtaCod = ""
    End If
    R.Close
            
    sSql = " SELECT PP.nPrdPersRelac,PP.cPersCod,C.cConsDescripcion, C.nConsValor"
    sSql = sSql & " from ProductoPersona PP"
    sSql = sSql & " Inner Join Constante C ON C.nConsValor = PP.nPrdPersRelac AND C.nConsCod = 3002"
    sSql = sSql & " Where cCtaCod = '" & sCtaCod & "' and nPrdPersRelac not in (20,28,29,21)"
            
    Set ObtenerOtrasRelaciones = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorObtenerOtrasRelaciones:
    Err.Raise Err.Number, "Obtener Otras Relaciones", Err.Description

End Function

Public Function CargarValoresPosicionCliente(ByRef pnLima As Integer, _
                                            ByRef pnSaldosAho As Integer)

Dim oDatos As COMDConstSistema.DCOMGeneral

On Error GoTo ErrorCargarValoresPosicionCliente

    Set oDatos = New COMDConstSistema.DCOMGeneral
    pnLima = oDatos.LeeConstSistema(103)
    
    pnSaldosAho = oDatos.LeeConstSistema(106)
    Set oDatos = Nothing
    
    Exit Function
ErrorCargarValoresPosicionCliente:
    Err.Raise Err.Number, "Valores Posicion Cliente", Err.Description

End Function

'Se modifico para manejar la Calificacion RCC
Public Function BuscarPosicionCliente(ByVal psPersCod As String, _
                                        ByVal pbMuestraAnalistas As Boolean, _
                                        ByVal pnLima As Integer, _
                                        ByVal pbPersNat As Boolean, _
                                        ByVal psPersDoc As String, _
                                        ByRef prsCred As ADODB.Recordset, _
                                        ByRef prsAho As ADODB.Recordset, _
                                        ByRef prsPig As ADODB.Recordset, _
                                        ByRef prsJud As ADODB.Recordset, _
                                        ByRef prsCF As ADODB.Recordset, _
                                        ByRef prsCom As ADODB.Recordset, _
                                        ByRef pdFechaRep As Date, _
                                        ByRef prsCalSBS As ADODB.Recordset, _
                                        ByRef prsEndSBS As ADODB.Recordset, _
                                        ByRef prsCalCMAC As ADODB.Recordset, _
                                        Optional pdFecSis As Date, _
                                        Optional nProm As Integer, Optional psPersTDoc As String = "1", _
                                        Optional ByRef prsDeuEnt As ADODB.Recordset) As Boolean 'FRHU 20140221 RQ14016 SE AGREGO prsDeuEnt) As Boolean

On Error GoTo ErrorBuscarPosicionCliente

BuscarPosicionCliente = True

Set prsCred = DatosPosicionCliente(psPersCod, pbMuestraAnalistas, pdFecSis, nProm)
Set prsAho = DatosPosicionClienteAhorro(psPersCod)
If pnLima = 1 Then           'Trujillo
    Set prsPig = DatosPosicionClientePigno(psPersCod)
Else 'Lima
    Set prsPig = DatosPosicionClientePignoLima(psPersCod)
End If
Set prsJud = DatosPosicionClienteJudicial(psPersCod)
Set prsCF = DatosPosicionClienteCF(psPersCod)
Set prsCom = DatosPosicionClienteComentarios(psPersCod)

'**** Calificacion RCC ****
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sCodDeudorRCC As String

sSql = " SELECT top 1 Fec_Rep FROM DBConsolidada..RCCTotal R  "
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
    BuscarPosicionCliente = False
    Exit Function
Else
    pdFechaRep = rs!Fec_Rep
End If
oCon.CierraConexion
Set oCon = Nothing

Set prsCalSBS = DatosPosicionClienteCalificacionSBS(pbPersNat, psPersDoc, sCodDeudorRCC, psPersTDoc)
If sCodDeudorRCC = "" Then
    Exit Function
End If
Set prsEndSBS = DatosPosicionClienteEndeudamientoSBS(sCodDeudorRCC)
Set prsCalCMAC = DatosPosicionClienteCalificacionCMAC(psPersCod)
Set prsDeuEnt = DatosPosicionClienteDeudaEnEntidades(psPersDoc, psPersTDoc) 'FRHU 20140221 RQ14016
'******************************
Exit Function

ErrorBuscarPosicionCliente:
    Err.Raise Err.Number, "Buscar Posicion Cliente", Err.Description

End Function

Public Function AdicionaRelacionesCredito(ByVal psPersCod As String, _
                                            ByRef pbTieneRelaciones As Boolean) As Variant

'ARCV 30-04-2007

'Dim oPers As COMDPersona.UCOMPersona
'Dim R As ADODB.Recordset
'Dim MatTemp() As String 'Variant
'Dim i As Integer
'
''ReDim MatTemp(0, 12)
'    pbTieneRelaciones = False
'
'    'Si es Titular Adicionamos el conyugue
'        Set oPers = New COMDPersona.UCOMPersona
'        Call oPers.ObtieneConyugeDePersona(psPersCod)
'        i = 0
'        If oPers.sPersCod <> "" Then
'            ReDim MatTemp(1, 12)
'            MatTemp(0, 0) = oPers.sPersCod
'            MatTemp(0, 1) = oPers.sPersNombre
'            MatTemp(0, 2) = "Conyuge"
'            MatTemp(0, 3) = Trim(Str(gColRelPersConyugue))
'            MatTemp(0, 4) = Trim(Str(gColRelPersConyugue))
'            MatTemp(0, 5) = oPers.sPersIdnroDNI
'            MatTemp(0, 6) = oPers.sPersIdnroRUC
'            MatTemp(0, 7) = NuevaRegistro
'            MatTemp(0, 8) = 1
'            MatTemp(0, 9) = oPers.sPersNatSexo
'            MatTemp(0, 10) = CInt(oPers.sPersPersoneria)
'            MatTemp(0, 11) = oPers.dPersNacCreac
'            i = i + 1
'            pbTieneRelaciones = True
'        End If
'        Set oPers = Nothing
'
'        'Obtenemos Otras Relaciones de Creditos Anteriores
'
'        Set R = ObtenerOtrasRelaciones(psPersCod)
'
'        If Not R.EOF Then
'            ReDim Preserve MatTemp(R.RecordCount + i, 12)
'            pbTieneRelaciones = True
'        End If
'
'        Do While Not R.EOF
'            Set oPers = New COMDPersona.UCOMPersona
'            Call oPers.ObtieneClientexCodigo(R!cPersCod)
'            If oPers.sPersCod <> "" Then
'                'ReDim Preserve MatTemp(i + 1, 12)
'                MatTemp(i, 0) = oPers.sPersCod
'                MatTemp(i, 1) = oPers.sPersNombre
'                MatTemp(i, 2) = Trim(R!cConsDescripcion)
'                MatTemp(i, 3) = Trim(Str(R!nPrdPersRelac))
'                MatTemp(i, 4) = Trim(Str(R!nPrdPersRelac))
'                MatTemp(i, 5) = oPers.sPersIdnroDNI
'                MatTemp(i, 6) = oPers.sPersIdnroRUC
'                MatTemp(i, 7) = NuevaRegistro
'                MatTemp(i, 8) = 1
'                MatTemp(i, 9) = oPers.sPersNatSexo
'                MatTemp(i, 10) = CInt(oPers.sPersPersoneria)
'                MatTemp(i, 11) = oPers.dPersNacCreac
'                i = i + 1
'            End If
'            Set oPers = Nothing
'            R.MoveNext
'        Loop
'
'        AdicionaRelacionesCredito = MatTemp

Dim oPers As COMDPersona.UCOMPersona
Dim R As ADODB.Recordset
Dim MatTemp() As String 'Variant
Dim i As Integer

'ReDim MatTemp(0, 12)
    pbTieneRelaciones = False
    
    'ARCV 20-06-2006
    
    'Si es Titular Adicionamos el conyugue
        Set oPers = New COMDPersona.UCOMPersona
        Call oPers.ObtieneConyugeDePersona(psPersCod)
        'ARCV 22-06-2006
        If oPers.sPersCod <> "" Then
            i = 1
        Else
            i = 0
        End If
        '-------------------
        Set R = ObtenerOtrasRelaciones(psPersCod)
        
        If Not R.EOF Then
            ReDim Preserve MatTemp(R.RecordCount + i, 12)
            pbTieneRelaciones = True
        Else
            ReDim MatTemp(i, 12)
        End If
                
        If oPers.sPersCod <> "" Then
            'ReDim MatTemp(1, 12)
            MatTemp(0, 0) = oPers.sPersCod
            MatTemp(0, 1) = oPers.sPersNombre
            MatTemp(0, 2) = "Conyuge"
            MatTemp(0, 3) = Trim(Str(gColRelPersConyugue))
            MatTemp(0, 4) = Trim(Str(gColRelPersConyugue))
            MatTemp(0, 5) = oPers.sPersIdnroDNI
            MatTemp(0, 6) = oPers.sPersIdnroRUC
            MatTemp(0, 7) = NuevaRegistro
            MatTemp(0, 8) = 1
            MatTemp(0, 9) = oPers.sPersNatSexo
            MatTemp(0, 10) = CInt(oPers.sPersPersoneria)
            MatTemp(0, 11) = oPers.dPersNacCreac
            i = 1
            pbTieneRelaciones = True
        End If
        Set oPers = Nothing
        
        'Obtenemos Otras Relaciones de Creditos Anteriores
            
'        Set R = ObtenerOtrasRelaciones(psPersCod)
'
'        If Not R.EOF Then
'            ReDim Preserve MatTemp(R.RecordCount + i, 12)
'            pbTieneRelaciones = True
'        End If
        
        Do While Not R.EOF
            Set oPers = New COMDPersona.UCOMPersona
            Call oPers.ObtieneClientexCodigo(R!cPersCod)
            If oPers.sPersCod <> "" Then
                'ReDim Preserve MatTemp(i + 1, 12)
                MatTemp(i, 0) = oPers.sPersCod
                MatTemp(i, 1) = oPers.sPersNombre
                MatTemp(i, 2) = Trim(R!cConsDescripcion)
                MatTemp(i, 3) = Trim(Str(R!nPrdPersRelac))
                MatTemp(i, 4) = Trim(Str(R!nPrdPersRelac))
                MatTemp(i, 5) = oPers.sPersIdnroDNI
                MatTemp(i, 6) = oPers.sPersIdnroRUC
                MatTemp(i, 7) = NuevaRegistro
                MatTemp(i, 8) = 1
                MatTemp(i, 9) = oPers.sPersNatSexo
                MatTemp(i, 10) = CInt(oPers.sPersPersoneria)
                MatTemp(i, 11) = oPers.dPersNacCreac
                i = i + 1
            End If
            Set oPers = Nothing
            R.MoveNext
        Loop
                
        AdicionaRelacionesCredito = MatTemp
End Function

Public Function ActualizaNivelesEnLote(ByVal pMatNiveles As Variant, _
                                        ByVal psPersCod As String)
Dim i As Integer

On Error GoTo ErrorActualizaNivelesEnLote

    For i = 0 To UBound(pMatNiveles) - 1
        If Mid(pMatNiveles(i), 1, 1) = "E" Then 'Elimina
            Call EliminarPermisoNivelApr(Mid(pMatNiveles(i), 2, Len(pMatNiveles(i)) - 1), psPersCod)
        Else    'Nuevo
            Call NuevoPermisoNivelApr(Mid(pMatNiveles(i), 2, Len(pMatNiveles(i)) - 1), psPersCod)
        End If
    Next
    
    Exit Function
ErrorActualizaNivelesEnLote:
    Err.Raise Err.Number, "Actualiza Niveles en Lote", Err.Description

End Function


'Cargar otras Relaciones de Personas Con un Credito
Public Function ObtenerMontoCheque(ByVal psDocNro As String) As Double

Dim sSql As String
Dim sCtaCod As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorObtenerOtrasRelaciones

    'Obtenemos Otras Relaciones de Creditos Anteriores
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        
    sSql = ""
    sSql = sSql & " SELECT ISNULL(sum(nMonto),0) Total FROM Mov M"
    sSql = sSql & " JOIN MovCol MC on MC.nMovNro=M.nMovNro"
    sSql = sSql & " Join MovDoc MD on MD.nMovNro=MC.nMovNro"
    sSql = sSql & " where cDocNro='" & psDocNro & "' and nNroCalen=3  and nMovFlag=0"
            
    Set R = oConecta.CargaRecordSet(sSql)
    
    If Not R.EOF And Not R.BOF Then
        ObtenerMontoCheque = R!Total
    Else
        ObtenerMontoCheque = 0
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorObtenerOtrasRelaciones:
    Err.Raise Err.Number, "No existe Cheque", Err.Description

End Function

'**DAOR 20070207
'**Recupera creditos que serán desembolsados en el banco de la nación
'**para generar el archivo de desembolso por corresponsalia
Public Function RecuperaCreditosParaDesemEnBcoNac(psFechaProceso As Date) As ADODB.Recordset
Dim sSql As String
Dim loConecta As COMConecta.DCOMConecta
    
    sSql = "Select  PR.cCtaCod,CB.cClave,substring(PR.cCtaCod,9,1) as Moneda, "
    sSql = sSql & " (case PT.nPersPersoneria when 1 then 1 else 2 end) as TipoPersona,PT.cPersNombre as NombreTitular, "
    sSql = sSql & " (case PT.nPersPersoneria "
    sSql = sSql & "  when 1  then "
    sSql = sSql & "     (select top 1 cast(PI1.cPersIDTpo as char(2))+PI1.cPersIDNro from PersId PI1 "
    sSql = sSql & "      where PI1.cPersCod=PT.cPersCod  and PI1.cPersIDTpo in (1,3,4,13) order by PI1.cPersIDTpo) "
    sSql = sSql & "  Else "
    sSql = sSql & "     (select cast(PI1.cPersIDTpo as char(2))+PI1.cPersIDNro from PersId PI1 "
    sSql = sSql & "      where PI1.cPersCod=PT.cPersCod  and PI1.cPersIDTpo=2) "
    sSql = sSql & "  end) as DocTitular, "
    sSql = sSql & " isnull(PRE.cPersNombre,'') as NombreRepresentante, "
    sSql = sSql & " isnull((select top 1 cast(PI2.cPersIDTpo as char(2))+PI2.cPersIDNro from PersId PI2 "
    sSql = sSql & "  where PI2.cPersCod=PRE.cPersCod  and PI2.cPersIDTpo in (1,3,4,13) order by PI2.cPersIDTpo),'') as DocRepresentante, "
    sSql = sSql & " CE.nMonto,CE.dPrdEstado,CE.nCuotas,CE.nColocCalendCod,CE.nPeriodoFechaFija,CE.nPeriodoGracia, "
    sSql = sSql & " CE.nPlazo,CB.cCodAgeBcoNac,PTI.nTasaInteres,right(PT.cPersCIIU,4) as SectorEconomico "
    sSql = sSql & " From    Producto PR inner join ProductoPersona PP on (PR.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20) "
    sSql = sSql & " inner join Persona PT on (PP.cPersCod=PT.cPersCod) "
    sSql = sSql & " inner join Colocaciones CO on (PR.cCtaCod=CO.cCtaCod) "
    sSql = sSql & " inner join ColocConvBcoNac CB on (CO.cCtaCod=CB.cCtaCod) "
    sSql = sSql & " inner join ColocacEstado CE on (CO.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002) "
    sSql = sSql & " inner join ProductoTasaInteres PTI on (PR.cCtaCod=PTI.cCtaCod and PTI.nPrdTasaInteres=1) "
    sSql = sSql & " left outer join ProductoPersona PP2 on (PR.cCtaCod=PP2.cCtaCod and PP2.nPrdPersRelac=23) "
    sSql = sSql & " left outer join Persona PRE on (PP2.cPersCod=PRE.cPersCod) "
    sSql = sSql & " Where  CB.nEstado=0 and CE.dPrdEstado <= '" & Format(psFechaProceso, "yyyymmdd") & " 23:59:59' "
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
    Set RecuperaCreditosParaDesemEnBcoNac = loConecta.CargaRecordSet(sSql)
        loConecta.CierraConexion
    Set loConecta = Nothing

End Function

'**DAOR 20070505
'**Recupera creditos para pagos por corresponsalia,
'**que serán enviados al banco de la nación
Public Function RecuperaCreditosParaPagosEnBcoNac(psFechaProceso As Date) As ADODB.Recordset
Dim sSql As String
Dim loConecta As COMConecta.DCOMConecta
    
    sSql = "Select  P.cCtaCod,PE.cPersNombre,C.nCuota,C.dVenc, "
    sSql = sSql & " sum(case CD.nPrdConceptoCod when 1000 then CD.nMonto-CD.nMontoPagado else 0 end) as nCapital, "
    sSql = sSql & " sum(case when CD.nPrdConceptoCod in(1100,1102) then CD.nMonto-CD.nMontoPagado else 0 end) as nInteres, "
    sSql = sSql & " sum(case CD.nPrdConceptoCod when 1101 then CD.nMonto-CD.nMontoPagado else 0 end) as nMora, "
    sSql = sSql & " sum(case when CD.nPrdConceptoCod>1200 and CD.nPrdConceptoCod<1300 then CD.nMonto-CD.nMontoPagado else 0 end) as nGastos "
    sSql = sSql & "From    Producto P inner join ColocacCred CC on P.cCtaCod=CC.cCtaCod "
    sSql = sSql & " inner join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20 "
    sSql = sSql & " inner join Persona PE on PP.cPersCod=PE.cPersCod "
    sSql = sSql & " inner join ColocCalendario C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen and CC.nNroProxCuota=C.nCuota "
    sSql = sSql & " inner join ColocCalendDet CD on C.cCtaCod=CD.cCtaCod and C.nCuota=CD.nCuota "
    sSql = sSql & " and C.nNroCalen=CD.nNroCalen and C.nColocCalendApl=CD.nColocCalendApl "
    sSql = sSql & "Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) and C.nColocCalendEstado = 0 and C.nColocCalendApl=1 "
    sSql = sSql & " and CC.cCtaCod not in (select cCtaCod from ColocConvBcoNac) "
'    sSql = sSql & " and substring(CC.cCtaCod, 6, 1) + Right(CC.cCtaCod, 11) not in (select CodBN from "
'    sSql = sSql & " (select substring(cCtaCod, 6, 1)+Right(cCtaCod, 11) as CodBN, count(*) as Veces "
'    sSql = sSql & " from Producto where nPrdEstado in (2020,2021,2022,2030,2031,2032)"
'    sSql = sSql & " group by substring(cCtaCod, 6, 1) + Right(cCtaCod, 11) having count(*)>1) T)"
    sSql = sSql & "Group by P.cCtaCod,PE.cPersNombre,C.nCuota,C.dVenc "
    sSql = sSql & "Order by P.cCtaCod,C.nCuota "
    
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
    Set RecuperaCreditosParaPagosEnBcoNac = loConecta.CargaRecordSet(sSql)
        loConecta.CierraConexion
    Set loConecta = Nothing

End Function

'**DAOR 20070505
'**Recupera creditos para pagos por corresponsalia,que serán enviados al
'**banco de la nación(Solo créditos pertenecientes al convenio del Bco. de la Nac.)
Public Function RecuperaCreditosParaPagosEnBcoNac_SoloConvBcoNac(psFechaProceso As Date) As ADODB.Recordset
Dim sSql As String
Dim loConecta As COMConecta.DCOMConecta
        
        sSql = "Select  P.cCtaCod,PE.cPersNombre,C.nCuota,C.dVenc,P.nSaldo, "
        sSql = sSql & " sum(case CD.nPrdConceptoCod when 1000 then CD.nMonto-CD.nMontoPagado else 0 end) as nCapital, "
        sSql = sSql & " sum(case when CD.nPrdConceptoCod in(1100,1102) then CD.nMonto-CD.nMontoPagado else 0 end) as nInteres, "
        sSql = sSql & " sum(case CD.nPrdConceptoCod when 1101 then CD.nMonto-CD.nMontoPagado else 0 end) as nMora, "
        sSql = sSql & " sum(case when CD.nPrdConceptoCod>1200 and CD.nPrdConceptoCod<1300 then CD.nMonto-CD.nMontoPagado else 0 end) as nGastos "
        sSql = sSql & "From    Producto P inner join ColocacCred CC on P.cCtaCod=CC.cCtaCod "
        sSql = sSql & " inner join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20 "
        sSql = sSql & " inner join Persona PE on PP.cPersCod=PE.cPersCod "
        sSql = sSql & " inner join ColocCalendario C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen "
        sSql = sSql & " inner join ColocCalendDet CD on C.cCtaCod=CD.cCtaCod and C.nCuota=CD.nCuota "
        sSql = sSql & " and C.nNroCalen=CD.nNroCalen and C.nColocCalendApl=CD.nColocCalendApl "
        sSql = sSql & " inner join ColocConvBcoNac CV on P.cCtaCod=CV.cCtaCod "
        sSql = sSql & "Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) and C.nColocCalendEstado = 0 and C.nColocCalendApl=1 "
        sSql = sSql & "Group by P.cCtaCod,PE.cPersNombre,C.nCuota,C.dVenc,P.nSaldo "
        sSql = sSql & "Order by P.cCtaCod,C.nCuota "
    
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
    Set RecuperaCreditosParaPagosEnBcoNac_SoloConvBcoNac = loConecta.CargaRecordSet(sSql)
        loConecta.CierraConexion
    Set loConecta = Nothing

End Function

'**DAOR 20070815
'**Función de que Recupera los registros para el reporte de saldos de cartera por analista,
'**obtenido de la base de datos consolidado
Public Function RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Integer, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double, Optional ByVal pbCondicBN As Boolean = False) As ADODB.Recordset 'DAOR 20070717, se agregó pnTipCam
    
Dim sSql As String
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadCondicionBN As String
Dim sCadProd As String
Dim oConecta As COMConecta.DCOMConecta


'*** PEAC 20080904
    sCadProd = ""
    If UBound(pMatProd) > 0 Then
        For i = 0 To UBound(pMatProd) - 1
            sCadProd = sCadProd & pMatProd(i) & ","
        Next i
        sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
    End If

'    sCadProd = " And SUBSTRING(Prd.cCtaCod,6,3) in ('"
'    For i = 0 To UBound(pMatProd) - 1
'        sCadProd = sCadProd & pMatProd(i) & "','"
'    Next i
'    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
'
'    If UBound(pMatProd) = 0 Then
'        sCadProd = ""
'    End If
    
'*** FIN PEAC


'*** PEAC 20080904
'    sCadAge = "('"
'    For i = 0 To UBound(pMatAgencias) - 1
'        sCadAge = sCadAge & pMatAgencias(i) & "','"
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
'*** FIN PEAC

    
'*** PEAC 20080904
    sCadCondicion = ""
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1)

'    sCadCondicion = "("
'    For i = 0 To UBound(pMatCondicion) - 1
'        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
'    Next i
'    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
'*** FIN PEAC
    
    '---Condición para ver solo creditos desemb en el BN GITU 05/04/2008---'
    If pbCondicBN Then
        '*** PEAC 20080905
        'sCadCondicionBN = " Join ColocConvBcoNac CBN On prd.cCtaCod = CBN.cCtaCod "
        sCadCondicionBN = "1" ''INCLUYE LA TABLA ColocConvBcoNac EN EL STP
    Else
        sCadCondicionBN = "2" ''NO INCLUYE LA TABLA ColocConvBcoNac EN EL STP
    End If
    
    '*** PEAC 20080308 - a este query se le agrego las columnas de num. clientes
'    sSql = "Select  "
'    sSql = sSql & " CASE WHEN GROUPING(A.cAgeDescripcion)=1 THEN 'TOTAL' ELSE ISNULL(A.cAgeDescripcion,'AGENCIA DESCONOCIDA') END  as cAgencia,"
'    sSql = sSql & " CASE WHEN GROUPING(Q.cUser)=1 THEN 'TOTAL' ELSE ISNULL(Q.cUser,'USUARIO DESCONOCIDA') END as cUser, "
'    sSql = sSql & " SUM(Q.nCar1Nro) as nCar1Nro, SUM(Q.nNumCli1) as nNumCli1,SUM(Q.nCar1Saldo) as nCar1Saldo,"
'    sSql = sSql & " SUM(Q.nCar2Nro) as nCar2Nro, SUM(Q.nNumCli2) as nNumCli2,SUM(Q.nCar2Saldo) as nCar2Saldo,"
'    sSql = sSql & " SUM(Q.nCar3Nro) as nCar3Nro, SUM(Q.nNumCli3) as nNumCli3,SUM(Q.nCar3Saldo) as nCar3Saldo,"
'    sSql = sSql & " SUM(Q.nCar4Nro) as nCar4Nro, SUM(Q.nNumCli4) as nNumCli4,SUM(Q.nCar4Saldo) as nCar4Saldo,"
'    sSql = sSql & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nNumCli5) as nNumCli5,SUM(Q.nCar5Saldo) as nCar5Saldo,"
'    sSql = sSql & " SUM(Q.nCar6Nro) as nCar6Nro, SUM(Q.nNumCli6) as nNumCli6,SUM(Q.nCar6Saldo) as nCar6Saldo"
'    sSql = sSql & " From ("
'    sSql = sSql & " Select cAgencia, T.cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nNumCli1) as nNumCli1,SUM(T.nCar1Saldo) as nCar1Saldo,"
'    sSql = sSql & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nNumCli2) as nNumCli2, SUM(T.nCar2Saldo) as nCar2Saldo,"
'    sSql = sSql & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nNumCli3) as nNumCli3, SUM(T.nCar3Saldo) as nCar3Saldo,"
'    sSql = sSql & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nNumCli4) as nNumCli4, SUM(T.nCar4Saldo) as nCar4Saldo,"
'    sSql = sSql & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nNumCli5) as nNumCli5, SUM(T.nCar5Saldo) as nCar5Saldo,"
'    sSql = sSql & "    SUM(T.nCar6Nro) as nCar6Nro, SUM(T.nNumCli6) as nNumCli6, SUM(T.nCar6Saldo) as nCar6Saldo"
'    sSql = sSql & " From ("
'    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, COUNT(*) as nCar1Nro, count(distinct pp.cperscod) as nNumCli1,SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar1Saldo, 0 as nCar2Nro,0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro,0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro,0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro,0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro,0 as nNumCli6, 0 as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    Where  Prd.nDiasAtraso >=" & P1I & " AND Prd.nDiasAtraso <=" & P1F
'    sSql = sSql & "        And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    sSql = sSql & "        And Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & "    Union "
'    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, 0 as nCar1Nro,0 as nNumCli1, 0 as nCar1Saldo, COUNT(*) as nCar2Nro,count(distinct pp.cperscod) as nNumCli2, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar2Saldo, 0 as nCar3Nro,0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro,0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nNumCli5,0 as nCar5Saldo, 0 as nCar6Nro,0 as nNumCli6, 0 as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    WHERE  Prd.nDiasAtraso >=" & P2I & " AND Prd.nDiasAtraso <=" & P2F
'    sSql = sSql & "        And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    sSql = sSql & "        And Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & "    Union "
'    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, 0 as nCar1Nro,0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro,0 as nNumCli2, 0 as nCar2Saldo, COUNT(*) as nCar3Nro,count(distinct pp.cperscod) as nNumCli3, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar3Saldo, 0 as nCar4Nro,0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro,0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro,0 as nNumCli6, 0 as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    WHERE Prd.nDiasAtraso >=" & P3I & " AND Prd.nDiasAtraso <=" & P3F
'    sSql = sSql & "        And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    sSql = sSql & "        And Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & "    Union "
'    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, 0 as nCar1Nro,0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro,0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro,0 as nNumCli3, 0 as nCar3Saldo, COUNT(*) as nCar4Nro,count(distinct pp.cperscod) as nNumCli4, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar4Saldo, 0 as nCar5Nro,0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro,0 as nNumCli6, 0 as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    WHERE Prd.nDiasAtraso > " & P4F
'    sSql = sSql & "        And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    sSql = sSql & "        And Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & "    Union "
'    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, 0 as nCar1Nro,0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro,0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro,0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro,0 as nNumCli4, 0 as nCar4Saldo, count(*) as nCar5Nro,count(distinct pp.cperscod) as nNumCli5, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar5Saldo, 0 as nCar6Nro,0 as nNumCli6, 0 as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    WHERE Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        And Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & "    Union "
'    sSql = sSql & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, Prd.cCodAnalista as cUser, 0 as nCar1Nro,0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro,0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro,0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro,0 as nNumCli4, 0 as nCar4Saldo , 0 as nCar5Nro,0 as nNumCli5, 0 as nCar5Saldo, count(*) as nCar6Nro, count(distinct pp.cperscod) as nNumCli6,SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldoCap else Prd.nSaldoCap*" & Str(pnTipCam) & " end) as nCar6Saldo "
'    sSql = sSql & "    From " & psServerCons & "CreditoConsol Prd "
'    sSql = sSql & "    INNER JOIN PRODUCTOPERSONA PP ON PRD.CCTACOD=PP.CCTACOD AND pp.nprdpersrelac=20"
'    sSql = sSql & sCadCondicionBN
'    sSql = sSql & "    WHERE Prd.nCondCre in " & sCadCondicion
'    sSql = sSql & "        And ((Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor & ") Or Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205))"
'    sSql = sSql & "        " & sCadProd
'    sSql = sSql & "    Group by Prd.cCodAnalista, substring(Prd.cCtaCod,4,2)"
'    sSql = sSql & " ) as T"
'    sSql = sSql & " Group by  T.cUser, cAgencia"
'    sSql = sSql & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
'    sSql = sSql & " WHERE Q.cAgencia IN " & sCadAge
'    sSql = sSql & " GROUP BY A.cAgeDescripcion, Q.cUser"
'    sSql = sSql & " With ROLLUP"
    
    '*** PEAC 20080905
    sSql = "exec stp_sel_RecuperaResumenSaldosCarteraPorAnalistaConsol " & pnTipCam & "," & P1I & "," & P1F & "," & P2I & "," & P2F & "," & P3I & "," & P3F & "," & P4F & ",'" & sCadCondicion & "','" & sCadProd & "','" & sCadAge & "','" & sCadCondicionBN & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalistaConsol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'*** PEAC 20080307
Public Function RecuperaResumenSaldosCarteraPorAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal cMatCondicion As String, ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double, Optional ByVal pbCondicBN As Boolean = False) As ADODB.Recordset 'DAOR 20070717, se agregó pnTipCam
    
    'ByVal pMatCondicion As Variant
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadCondicionBN As String
Dim sCadProd As String
Dim sAnalistas As String
'Dim oGen As DGeneral
Dim oGen As COMDConstSistema.DCOMGeneral

    'Set oGen = New DGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    Set oGen = Nothing

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    '---Condición para ver solo creditos desemb en el BN GITU 05/04/2008---'
    If pbCondicBN Then
        sCadCondicionBN = " Join ColocConvBcoNac CBN On prd.cCtaCod = CBN.cCtaCod "
    End If
    
    '**DAOR 20070717**********************************************************
    sSql = "Select  "
    sSql = sSql & " CASE WHEN GROUPING(A.cAgeDescripcion)=1 THEN 'TOTAL'"
    sSql = sSql & "   ELSE ISNULL(A.cAgeDescripcion,'AGENCIA DESCONOCIDA')"
    sSql = sSql & " END  as cAgencia,"
    sSql = sSql & " CASE WHEN GROUPING(Q.cUser)=1 THEN 'TOTAL'"
    sSql = sSql & "   ELSE ISNULL(Q.cUser,'AGENCIA DESCONOCIDA')"
    sSql = sSql & " END as cUser,"
    sSql = sSql & " SUM(Q.nCar1Nro) as nCar1Nro, SUM(Q.nNumCli1) as nNumCli1, SUM(Q.nCar1Saldo) as nCar1Saldo,"
    sSql = sSql & " SUM(Q.nCar2Nro) as nCar2Nro, SUM(Q.nNumCli2) as nNumCli2, SUM(Q.nCar2Saldo) as nCar2Saldo,"
    sSql = sSql & " SUM(Q.nCar3Nro) as nCar3Nro, SUM(Q.nNumCli3) as nNumCli3, SUM(Q.nCar3Saldo) as nCar3Saldo,"
    sSql = sSql & " SUM(Q.nCar4Nro) as nCar4Nro, SUM(Q.nNumCli4) as nNumCli4, SUM(Q.nCar4Saldo) as nCar4Saldo,"
    sSql = sSql & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nNumCli5) as nNumCli5, SUM(Q.nCar5Saldo) as nCar5Saldo,"
    sSql = sSql & " SUM(Q.nCar6Nro) as nCar6Nro, SUM(Q.nNumCli6) as nNumCli6, SUM(Q.nCar6Saldo) as nCar6Saldo"
    sSql = sSql & " From ("
    sSql = sSql & " Select cAgencia, T.cUser + ' ' +  T.cPersNombre as cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nNumCli1) as nNumCli1, SUM(T.nCar1Saldo) as nCar1Saldo,"
    sSql = sSql & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nNumCli2) as nNumCli2, SUM(T.nCar2Saldo) as nCar2Saldo,"
    sSql = sSql & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nNumCli3) as nNumCli3, SUM(T.nCar3Saldo) as nCar3Saldo,"
    sSql = sSql & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nNumCli4) as nNumCli4, SUM(T.nCar4Saldo) as nCar4Saldo,"
    sSql = sSql & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nNumCli5) as nNumCli5, SUM(T.nCar5Saldo) as nCar5Saldo,"
    sSql = sSql & "    SUM(T.nCar6Nro) as nCar6Nro, SUM(T.nNumCli6) as nNumCli6, SUM(T.nCar6Saldo) as nCar6Saldo"
    sSql = sSql & " From ("
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, COUNT(*) as nCar1Nro, count(distinct pp.cperscod) as nNumCli1, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar1Saldo, 0 as nCar2Nro, 0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nNumCli6, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "    Where  CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nNumCli1, 0 as nCar1Saldo, COUNT(*) as nCar2Nro, count(distinct pp.cperscod) as nNumCli2, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar2Saldo, 0 as nCar3Nro, 0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nNumCli6, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "     WHERE  CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nNumCli2, 0 as nCar2Saldo, COUNT(*) as nCar3Nro, count(distinct pp.cperscod) as nNumCli3, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar3Saldo, 0 as nCar4Nro, 0 as nNumCli4, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nNumCli6, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "     WHERE CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nNumCli3, 0 as nCar3Saldo, COUNT(*) as nCar4Nro, count(distinct pp.cperscod) as nNumCli4, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar4Saldo, 0 as nCar5Nro, 0 as nNumCli5, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nNumCli6, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "     WHERE CC.nDiasAtraso > " & P4F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nNumCli4, 0 as nCar4Saldo, count(*) as nCar5Nro, count(distinct pp.cperscod) as nNumCli5, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar5Saldo, 0 as nCar6Nro, 0 as nNumCli6, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nNumCli1, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nNumCli2, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nNumCli3, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nNumCli4, 0 as nCar4Saldo , 0 as nCar5Nro, 0 as nNumCli5, 0 as nCar5Saldo, count(*) as nCar6Nro, count(distinct pp.cperscod) as nNumCli6, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And ((Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor & ") Or Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205))"
    sSql = sSql & "        Inner Join ProductoPersona PP ON PRD.cctacod=pp.cctacod and pp.nprdpersrelac=20"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & sCadCondicionBN
    sSql = sSql & "     WHERE CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & " ) as T"
    sSql = sSql & " Group by  T.cUser + ' ' +  T.cPersNombre, cAgencia"
    sSql = sSql & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
    sSql = sSql & " WHERE Q.cAgencia IN " & sCadAge
    sSql = sSql & " GROUP BY A.cAgeDescripcion, Q.cUser"
    sSql = sSql & " With ROLLUP"
    '***************************************************************************
    
    'ssql = ssql & "     WHERE CC.nColocCondicion in " & sCadCondicion
    
    Set oConecta = New COMConecta.DCOMConecta 'DConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function




'**DAOR 20070815
'**Función de que Recupera los registros para el reporte consolidado de cartera por analista,
'**obtenido de la base de datos consolidado
Public Function RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double) As ADODB.Recordset 'DAOR 20070717, se agregó pnTipCam
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadProd As String

    '****** BRGO - BASILEA II
    sCadProd = " And CCT.cTpoCredCod in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSql = "Select  "
    sSql = sSql & " case when grouping(A.cAgeDescripcion)=1 then 'TOTAL' else isnull(A.cAgeDescripcion,'AGENCIA DESCONOCIDA') end  as cAgencia, "
    sSql = sSql & " case when grouping(CC.cCodAnalista)=1 then 'TOTAL' else isnull(CC.cCodAnalista ,'ANALISTA DESCONOCIDO') END as cUser, "
    sSql = sSql & "   sum(case CC.nCondCre when 1 then 1 else 0 end) as nCantNuevo, "
    sSql = sSql & "   sum(case CC.nCondCre when 1 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoNuevo, "
    sSql = sSql & "   sum(case CC.nCondCre when 2 then 1 else 0 end) as nCantRecurrente, "
    sSql = sSql & "   sum(case CC.nCondCre when 2 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoRecurrente, "
    sSql = sSql & "   sum(case CC.nCondCre when 3 then 1 else 0 end) as nCantParalelo, "
    sSql = sSql & "   sum(case CC.nCondCre when 3 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoParalelo, "
    sSql = sSql & "   sum(case CC.nCondCre when 4 then 1 else 0 end) as nCantRefinanciado, "
    sSql = sSql & "   sum(case CC.nCondCre when 4 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoRefinanciado, "
    sSql = sSql & "   sum(case CC.nCondCre when 5 then 1 else 0 end) as nCantAmpliado, "
    sSql = sSql & "   sum(case CC.nCondCre when 5 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoAmpliado, "
    sSql = sSql & "   sum(case CC.nCondCre when 6 then 1 else 0 end) as nCantAutomatico, "
    sSql = sSql & "   sum(case CC.nCondCre when 6 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoAutomatico, "
    sSql = sSql & "   sum(case CC.nCondCre when 7 then 1 else 0 end) as nCantAdicional, "
    sSql = sSql & "   sum(case CC.nCondCre when 7 then case substring(CC.cCtaCod,9,1) when '1' then CC.nMontoDesemb else CC.nMontoDesemb*" & Str(pnTipCam) & "  end else 0 end) as nMontoAdicional "
    sSql = sSql & " From " & psServerCons & "CreditoConsol CC "
    sSql = sSql & "   Inner Join " & psServerCons & "CreditoConsolTotal CCT ON CC.cCtaCod = CCT.cCtaCod "
    sSql = sSql & "   Inner Join Agencias A ON CCT.cAgeCodAct=A.cAgeCod "
    sSql = sSql & " where CCT.cAgeCodAct IN " & sCadAge
    sSql = sSql & "   and ((CC.nPrdEstado >= " & gColocEstVigNorm & " and CC.nPrdEstado <= " & gColocEstRefMor & ") Or CC.nPrdEstado in (" & gColocEstRecVigJud & ",2205))" 'Vigentes y judiciales
    sSql = sSql & "   and CC.dFecVig between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & " group by A.cAgeDescripcion,CC.cCodAnalista with rollup "
    '**************************
    'JUEZ 20130603 Se agregó case nCondCre 7
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaConsolidadoCarteraPorAnalistaConsol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'**DAOR 20071210
'**Función de que Recupera los registros para el reporte de seguro de desgravamen,
'**obtenido de la base de datos consolidado
Public Function RecuperaSeguroDesgravamenConsol(ByVal pdFechaProc As Date, ByVal pnTipCam As Double, Optional ByVal pdFechaProc01 As Date, Optional ByVal pnMoneda As String = "1", Optional ByVal pnInd As Boolean = "false", Optional ByVal pnTipInd As Boolean = "false", Optional ByVal pnTipIndCartera As Boolean = "false") As ADODB.Recordset

Dim sSql As String
Dim sMonedasT As String
Dim oConecta As COMConecta.DCOMConecta

    If Len(Trim(pnMoneda)) = 0 Then
        sMonedasT = "1"
    Else
        sMonedasT = Replace(Replace(pnMoneda, "'", ""), " ", "")
    End If
    
    If pnTipIndCartera = False Then
        If pnTipInd Then
            If pnInd Then
                sSql = " exec stp_sel_SeguroDesgravamenTitu " & pnTipCam & ",'" & Format(pdFechaProc01, "yyyymmdd") & "','" & Format(pdFechaProc, "yyyymmdd") & "','" & sMonedasT & "' "
            Else
                sSql = " exec stp_sel_SeguroDesgravamenManco " & pnTipCam & ",'" & Format(pdFechaProc01, "yyyymmdd") & "','" & Format(pdFechaProc, "yyyymmdd") & "','" & sMonedasT & "' "
            End If
        Else
            If pnInd Then
                sSql = " exec stp_sel_SeguroDesgravamenTituCancelados " & pnTipCam & ",'" & Format(pdFechaProc01, "yyyymmdd") & "','" & Format(pdFechaProc, "yyyymmdd") & "','" & sMonedasT & "' "
            Else
                sSql = " exec stp_sel_SeguroDesgravamenMancoCancelados " & pnTipCam & ",'" & Format(pdFechaProc01, "yyyymmdd") & "','" & Format(pdFechaProc, "yyyymmdd") & "','" & sMonedasT & "' "
            End If
        End If
    Else
        sSql = " exec stp_sel_SeguroDesgravamenCartera " & pnTipCam & ",'" & Format(pdFechaProc, "yyyymmdd") & "' "
    End If
    
    'sSql = " exec stp_sel_ReporteSeguroDesgravamen " & pnTipCam & ",'" & Format(pdFechaProc, "yyyymmdd") & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSeguroDesgravamenConsol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'**ALPA 20080404
'**Función de que Recupera los registros para el reporte de creditos cancelados x ampliacion
'**obtenido de la base de datos consolidado
Public Function RecuperaCredCanceXAmpliacion(ByVal pdFechaIn As Date, ByVal pdFechaFi As Date, Optional sAgeCod As Variant) As ADODB.Recordset
   
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    '**20080606, comentado por error al compilar *************************************
    'sSql = " exec stp_sel_ReporteSeguroDesgravamen " & pnTipCam & ",'" & Format(pdFechaProc, "yyyymmdd") & "' "
    '*******************************************************************
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredCanceXAmpliacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'madm pago convenio bn--------------------------------------------------------------------
'madm 20100201-----------------------------------------------------------------------
Public Function CargaDatos_eliminar(ByVal fProceso As String) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim sSqlaux As String

    On Error GoTo ERRORCargaDatos_eliminar
    
    sSql = "select id,cCodBco,cCodCli,nNumReg nImpToTS,nImpToTD,fProceso, tRegistro"
    sSql = sSql & " from LogCobrosBcoNac where convert(varchar(12),fProceso,103)='" & fProceso & "'"
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set CargaDatos_eliminar = R
    Set R = Nothing
    Exit Function
ERRORCargaDatos_eliminar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'cambio  xstore 20111107 - 20100616
Public Function CargaDatos_Cabecera_108121(ByVal nImpToTS As Currency, ByVal nImpToTD As Currency, ByVal nNumReg As Integer, ByVal fProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacCabecera " & nImpToTS & "," & nImpToTD & "," & nNumReg & " , '" & fProceso & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaDatos_Cabecera_108121 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function DevolverDatos_Cabecera_108121(ByVal id As Double, Optional x As Boolean = False) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim sSqlaux As String

If x Then
    sSqlaux = " AND est = 1 "
End If

    On Error GoTo ERRORCargaDatos_Comite_ult
    
    sSql = " select id,cCodBco,cCodCli,nNumReg , isnull(nImpToTS,0) nImpToTS ,   isnull(nImpToTD,0) nImpToTD , "
    sSql = sSql & " nImpToTS1 = isnull((select sum(nImporteCobrado) from LogCobrosBcoNacDet where est = 0 and nMoneda = 1 and Id_cab =  " & id & " ),0) , "
    sSql = sSql & " nImpToTD1 = isnull((select sum(nImporteCobrado) from LogCobrosBcoNacDet where est = 0 and nMoneda <> 1 and Id_cab =  " & id & " ),0) , "
    sSql = sSql & " fProceso, tRegistro from LogCobrosBcoNac where Id = " & id & " and ext = 0"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set DevolverDatos_Cabecera_108121 = R
    Set R = Nothing
    Exit Function
ERRORCargaDatos_Comite_ult:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

'cambio  xstore 20100616
Public Function DevolverDatos_Cabecera_108121_NoProcesados(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalleNoPagados " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_108121_NoProcesados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'cambio  xstore 20100616
Public Function DevolverDatos_previoCabecera_108121(ByVal pcCtaCod As String, ByVal pNomCli As String, Optional Filtro1 As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacCtaNombre '" & pcCtaCod & "','" & pNomCli & "'," & IIf(Filtro1, 1, 0) & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_previoCabecera_108121 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'cambio  xstore 20100616
Public Function DevolverDatos_Detalle_108121(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalle " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_108121 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'cambio  xstore 20100616
Public Function DevolverDatos_Detalle_108121_NoPagados(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalleNo " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_108121_NoPagados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'end madm
'madm 20111019
Public Sub Insertar_Datos_Cabecera_BN(ByVal prs As ADODB.Recordset, Optional ByVal nValor As Integer = 0)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Cabecera_BN
   
   sSql = "Exec stp_ins_LogCobrosBcoNac '" & prs("cCodBco") & "','" & prs("cCodCli") & "'," & prs("nNumReg") & "," & prs("nImpToTS") & "," & prs("nImpToTD") & ", '" & prs("fProceso") & "', '" & prs("tRegistro") & "', " & nValor & ""
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertar_Datos_Cabecera_BN:
    Err.Raise Err.Number, "Insertar Datos Cabecera", Err.Description
End Sub
'madm 20111109 - 6
Public Sub Actualizar_DetallexEst_BN(ByVal id As Double, ByVal cCodCtaF As String, ByVal nNroCta As String, ByVal dFechaCobro As String, ByVal nTipoPago As Integer, Optional ByVal bValorAct As Boolean = True)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim nInt As Integer
    On Error GoTo ErrorActualizar_DetallexEst_BN
    
   nInt = IIf(bValorAct, 1, 0)
   
   If nTipoPago = 1 Then
        sSql = "Exec stp_upd_LogCobrosBcoNacDet " & id & ",'" & cCodCtaF & "','" & nNroCta & "','" & dFechaCobro & "' ," & nInt & "  "
   Else
        sSql = "Exec stp_upd_LogCorresponsaliaBcoNacDet " & id & ",'" & cCodCtaF & "','" & nNroCta & "','" & dFechaCobro & "' ," & nInt & "  "
   End If
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizar_DetallexEst_BN:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub

Public Sub Insertar_Datos_Detalle_BN(ByVal id As Double, ByVal prs As ADODB.Recordset, Optional x As Boolean = False, _
                                     Optional ByVal pConnBase As COMConecta.DCOMConecta = Nothing)
                                     'RIRO20170328 ADD pConnBase
    'oConecta.AbreConexion
    Dim sSql As String
    Dim aux As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Detalle_BN
   
   If x Then
    aux = 1
   Else
    aux = 0
   End If
   
   sSql = "Exec stp_ins_LogCobrosBcoNacDet " & id & ",'" & prs("cCodCta") & "','" & prs("cCodCtaF") & "','" & prs("nNroCta") & "','" & prs("cSituacion") & "','" & prs("nmoneda") & "','" & prs("cNomCliente") & "', '" & prs("nImporteCuota") & "', '"
   sSql = sSql & prs("dCaducidad") & "','" & prs("nIndTasa") & "', '"
   sSql = sSql & prs("nfactorMora") & "','" & prs("nFactorCompensacion") & "','" & prs("nImporteGastos") & "' ,'"
   sSql = sSql & prs("cCuentaCliente") & "','" & prs("cOrdenCobro") & "','" & prs("nMora") & "','" & prs("nCompensacion") & "', '"
   sSql = sSql & prs("nImporteCobrado") & "','" & prs("cAgencia") & "','" & prs("dFechaCobro") & "' ,'"
   sSql = sSql & prs("dHoraCobro") & "' , '00',"
   sSql = sSql & aux
       
   If Not pConnBase Is Nothing Then
    Set oConecta = pConnBase
   Else
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
   End If
       
    oConecta.ConexionActiva.Execute sSql
       
   If pConnBase Is Nothing Then
    oConecta.CierraConexion
    Set oConecta = Nothing
   End If
   
    Exit Sub
ErrorInsertar_Datos_Detalle_BN:
    Err.Raise Err.Number, "Error Insertar Datos Detalle BN", Err.Description
End Sub

'madm 20100225
'RIRO20170324 add'pnConnBase
Public Sub Insertar_Datos_Sobrantes(ByVal cCtaCod As String, ByVal nCuota As Integer, ByVal nMontoDif As Double, ByVal dfpago As Date, ByVal cNomCliente As String, ByVal nMovNro As Long, _
                                    Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Sobrantes
   
   sSql = "Exec stp_ins_MovCobrosBcoNacSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
       
       If Not pnConnBase Is Nothing Then
        Set oConecta = pnConnBase
       Else
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
       End If
       oConecta.ConexionActiva.Execute sSql
       
       If pnConnBase Is Nothing Then
        oConecta.CierraConexion
        Set oConecta = Nothing
       End If
    Exit Sub
ErrorInsertar_Datos_Sobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub
'20111014
Public Sub Eliminar_Pago_Cab_Det(ByVal id As Integer, ByVal tipo As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Sobrantes
   
    sSql = "Exec stp_del_LogCobrosBcoNac " & id & " , " & tipo & ""
   
   Set oConecta = New COMConecta.DCOMConecta
   oConecta.AbreConexion
   oConecta.ConexionActiva.Execute sSql
   oConecta.CierraConexion
   Set oConecta = Nothing
Exit Sub
ErrorInsertar_Datos_Sobrantes:
    Err.Raise Err.Number, "Eliminar Datos Generales", Err.Description
End Sub
'madm 20100316 ---------------------------------------------------------------
Public Function ReportesListaRecuperacionesBN(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_lista_recuperaciones_bn '" & Format(pdFecha, "yyyymmdd") & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesListaRecuperacionesBN = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesListaRecuperacionesBN = Null
End Function

Public Function ReportesListaCorresponsaliaBN(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_lista_corresponsalia_bn '" & Format(pdFecha, "yyyymmdd") & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesListaCorresponsaliaBN = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesListaCorresponsaliaBN = Null
End Function
'Pago bn
Public Function DevolverIDDatos_Detalle(ByVal id_cab As Double, ByVal nCuota As Integer, ByVal cCtaCod As String, ByVal dfpago As String, _
                                         Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing) As ADODB.Recordset
                                         'RIRO20170328 ADD pnConnBase
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalleID " & id_cab & " , " & nCuota & " , '" & cCtaCod & "' , '" & dfpago & "'"
    
    If Not pnConnBase Is Nothing Then
        Set oConecta = pnConnBase
    Else
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
    End If
    
    Set DevolverIDDatos_Detalle = oConecta.CargaRecordSet(sSql)
    
    If pnConnBase Is Nothing Then
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
    
End Function
'-----------------------------------------------------------------------------
'MADM CORRESPONSALIA - 20111109 - 1
Public Function CargaDatos_Cabecera_corresponsalia(ByVal nImpToTS As Currency, ByVal nNumReg As Integer, ByVal fProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacCabecera " & nImpToTS & "," & nNumReg & " , '" & fProceso & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaDatos_Cabecera_corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub Insertar_Datos_SobrantesCorresponsalia(cCtaCod As String, nCuota As Integer, nMontoDif As Double, dfpago As Date, cNomCliente As String, nMovNro As Long, _
                                                  Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing)
    'RIRO 20170612 ADD pnConnBase
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Sobrantes
   
   sSql = "Exec stp_ins_MovCorresponsaliaBcoNacSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
   
    If Not pnConnBase Is Nothing Then
        Set oConecta = pnConnBase
    Else
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
    End If

    oConecta.ConexionActiva.Execute sSql
    
    If pnConnBase Is Nothing Then
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
       
    Exit Sub
ErrorInsertar_Datos_Sobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub
'20111109 - 3
Public Sub Insertar_Datos_Detalle_corresponsaliaBN(ByVal id As Double, ByVal prs As ADODB.Recordset, Optional x As Boolean = False, Optional ByVal pConnBase As COMConecta.DCOMConecta = Nothing)
    
    'RIRO20170613 ADD pConnBase
    
    Dim sSql As String
    Dim aux As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Detalle_corresponsaliaBN
   If x Then
    aux = 1
   Else
    aux = 0
   End If
    
   sSql = "Exec stp_ins_LogCorresponsaliaBcoNacDet " & id & ",'" & prs("cCodCta") & "','" & prs("cCodCtaF") & "','" & prs("nNroCta") & "','" & prs("nmoneda") & "','" & prs("cNomCliente") & "', '" & prs("nImporteCobrado") & "', '"
   sSql = sSql & prs("dFechaCobro") & "',"
   sSql = sSql & aux

    If Not pConnBase Is Nothing Then
     Set oConecta = pConnBase
    Else
     Set oConecta = New COMConecta.DCOMConecta
     oConecta.AbreConexion
    End If
    
    oConecta.ConexionActiva.Execute sSql
       
    If pConnBase Is Nothing Then
     oConecta.CierraConexion
     Set oConecta = Nothing
    End If

    Exit Sub
ErrorInsertar_Datos_Detalle_corresponsaliaBN:
    Err.Raise Err.Number, "Error Insertar Datos Detalle corresponsalia BN", Err.Description
End Sub
'20111109 - 4
Public Function DevolverIDDatos_DetalleCorresponsalia(ByVal id_cab As Double, ByVal nCuota As Integer, ByVal cCtaCod As String, ByVal dfpago As String, _
                                                    Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing) As ADODB.Recordset
                                                    'RIRO20170328 ADD pnConnBase
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalleID " & id_cab & " , " & nCuota & " , '" & cCtaCod & "' , '" & dfpago & "'"

    If Not pnConnBase Is Nothing Then
        Set oConecta = pnConnBase
    Else
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
    End If

    Set DevolverIDDatos_DetalleCorresponsalia = oConecta.CargaRecordSet(sSql)
    
    If pnConnBase Is Nothing Then
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
End Function
'pago bn cobros
Public Sub Actualizar_DetallexEst_Mov(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, ByVal nMovNroC As String, _
                                      Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing)
    'RIRO20170324 pnConnBase

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_DetallexEst_Mov
    
    sSql = "Exec stp_upd_LogCobrosBcoNacDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'"

    If Not pnConnBase Is Nothing Then
        Set oConecta = pnConnBase
    Else
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
    End If

    oConecta.ConexionActiva.Execute sSql

    If pnConnBase Is Nothing Then
       oConecta.CierraConexion
       Set oConecta = Nothing
    End If

    Exit Sub
ErrorActualizar_DetallexEst_Mov:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub

Public Sub Actualizar_DetallexEst_MovCorresponsalia(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, ByVal nMovNroC As String, Optional ByVal nComision As Double, _
                                                    Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing)
    'RIRO20170324 ADD pnConnBase
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_DetallexEst_Mov
 
   sSql = "Exec stp_upd_LogCorresponsaliaBcoNacDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'," & nComision & ""
   
   If Not pnConnBase Is Nothing Then
    Set oConecta = pnConnBase
   Else
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
   End If
   
   oConecta.ConexionActiva.Execute sSql
   
   If pnConnBase Is Nothing Then
       oConecta.CierraConexion
       Set oConecta = Nothing
   End If
      
   Exit Sub
ErrorActualizar_DetallexEst_Mov:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub

Public Sub Actualizar_DetallexEst_CorresponsaliaBN(ByVal id As Double, ByVal cCodCtaF As String, ByVal nNroCta As String, ByVal dFechaCobro As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_DetallexEst_BN

   sSql = "Exec stp_upd_LogCorresponsaliaBcoNacDet " & id & ",'" & cCodCtaF & "','" & nNroCta & "','" & dFechaCobro & "' ,1 "
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorActualizar_DetallexEst_BN:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
'20111109 - 2
Public Sub Insertar_Datos_Cabecera_corresponsaliaBN(ByVal prs As ADODB.Recordset, Optional ByVal nValor As Integer = 0)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Cabecera_corresponsaliaBN
   
   sSql = "Exec stp_ins_LogCorresponsaliaBcoNac " & prs("nNumReg") & "," & prs("nImpToTS") & ",'" & prs("fProceso") & "'," & nValor & ""
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertar_Datos_Cabecera_corresponsaliaBN:
    Err.Raise Err.Number, "Insertar Datos Cabecera", Err.Description
End Sub

Public Function DevolverDatos_Cliente_Corresponsalia(ByVal pcCtaCod As String, ByVal pNomCli As String, Optional Filtro1 As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacCtaNombre '" & pcCtaCod & "','" & pNomCli & "'," & IIf(Filtro1, 1, 0) & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cliente_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'reporte bn
Public Function DevolverDatos_Cabecera_reporte(ByVal dfpago As Date, Optional est As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacReporte  '" & Format(dfpago, "YYYYMMDD") & "'," & est & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_reporte = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'20111109 - reporte
Public Function DevolverDatos_Cabecera_reporte_Corresponsalia(ByVal dfpago As Date, Optional est As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacReporte  '" & Format(dfpago, "YYYYMMDD") & "'," & est & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_reporte_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'MADM 20111116
Public Function DevolverDatos_Cabecera_reporte_CorresponsaliaPago(ByVal id As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacReportePago  " & id & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_reporte_CorresponsaliaPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevolverDatos_Detalle_108121_Corresponsalia(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalle " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_108121_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DevolverDatos_Cabecera_108121_NoProcesados_Corresponsalia(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalleNoPagados " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_108121_NoProcesados_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DevolverDatos_Detalle_108121_NoPagados_Corresponsalia(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalleNo " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_108121_NoPagados_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'reporte BN
Public Function DevolverDatos_Cabecera_Sobrante(ByVal dfpago As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacCabeceraSobrante " & Format(dfpago, "yyyymmdd") & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_Sobrante = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevolverDatos_Cabecera_Sobrante_Corresponsalia(ByVal dfpago As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacCabeceraSobrante " & Format(dfpago, "yyyymmdd") & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_Sobrante_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevolverDatos_Detalle_Sobrante(ByVal dfpago As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalleSobrante '" & Format(dfpago, "yyyymmdd") & "'  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_Sobrante = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevolverDatos_Detalle_Sobrante_Corresponsalia(ByVal dfpago As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalleSobrante '" & Format(dfpago, "yyyymmdd") & "'  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Detalle_Sobrante_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'MADM 20100626 - FINANCIERO
Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim SQL As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            SQL = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM " & vsServerCom & "Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            SQL = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM " & vsServerCom & "AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           SQL = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
                & "             CASE " _
                & "                  WHEN CF.cSubctaCod IS NULL THEN AA.cSubCtaCod " _
                & "                  Else CF.cSubCtaCod END As SubCta " _
                & "    FROM     " & vsServerCom & "AREAAGENCIA AA " _
                & "             LEFT JOIN ( Select cAreaCod , cAgeCod , cSubCtaCod " _
                & "                          From " & vsServerCom & "CtaAreaAgeFiltro " _
                & "                          WHERE  cCtaContcod in ('" & psCtaContCod & "')) AS CF " _
                & "             ON CF.cAreacod = AA.cAreaCod and CF.cAgeCod = AA.cAgeCod " _
                & "    WHERE    AA.cAgeCod='" & Mid(psObjetoCod, 4, 2) & "' AND  AA.cAreaCod='" & Mid(psObjetoCod, 1, 3) & "'"
                
        Case ObjEntidadesFinancieras
            'If Len(psObjetoCod) > 15 Then
                lsTpoIf = Mid(psObjetoCod, 1, 2)
                lsPersCtaIF = Mid(psObjetoCod, 4, 13)
                lsCtaIf = Mid(psObjetoCod, 18, 10)
            'Else
            '    If psPersCodIf <> "" Then Exit Function
            '    lsPersCtaIF = psPersCodIf
            '    lsCtaIf = psObjetoCod
            'End If
            If Len(psObjetoCod) = 13 Then
                SQL = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                SQL = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            SQL = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oConect.CargaRecordSet(SQL)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    'Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function


Public Function GrabaDeposiBancoNegocioFinan(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                        ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                        ByVal psDocNroVoucher As String, _
                                        ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                        Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psctaAho As String = "", _
                                        Optional ByVal pnDocTpoNANCAux As TpoDoc = -1, Optional ByVal psNroDocNANCAux As String = "", Optional ByVal pnImporteAux As Currency, _
                                        Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = "", _
                                        Optional ByVal pbGrabaNegocio As Boolean = False, Optional ByVal pbBitCentral As Boolean = False, _
                                        Optional psOpeCodNegocio As String = "", _
                                        Optional psAjusteCta As String = "", Optional pnAjusteMonto As Currency = 0, Optional pnTCBanco As Currency = 0, Optional pnMovNroRef As Long, _
                                        Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing) As Integer

'RIRO 20170612 ADD pnConnBase

Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As COMDMov.DCOMMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lbTransGeneral As Boolean
Dim lsSubCta As String
Set oDMov = New COMDMov.DCOMMov
Dim lsMovNroNegocio As String
Dim lsMsgErr As String

On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDeposiBancoNegocioFinan = 1
lbTransGeneral = False
lbTrans = False

If Not pnConnBase Is Nothing Then
    Call oDMov.SetConexion(pnConnBase)
    oDMov.bTransaccion = True
    lbTransGeneral = True
End If

If Not lbTransGeneral Then
    oDMov.BeginTrans
    lbTrans = True
End If

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
lsSubCta = GetFiltroObjetos(ObjEntidadesFinancieras, psCtaContBanco, psCtaIf, False)

oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContBanco + lsSubCta, pnImporte * IIf(pbDeposito, 1, -1)

lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psCtaIf, 4, 13), Mid(psCtaIf, 1, 2), Mid(psCtaIf, 18, 10)

lnMovItem = lnMovItem + 1: lnMovOrden = 0

oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, (pnImporte - pnAjusteMonto) * IIf(pbDeposito, -1, 1)

oDMov.InsertaMovRef lnMovNro, pnMovNroRef

GrabaDeposiBancoNegocioFinan = 0

If Mid(psOpeCod, 3, 1) = 2 Then
    oDMov.GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

If Not lbTransGeneral Then
    oDMov.CommitTrans
    lbTrans = False
    Set oDMov = Nothing
End If


Exit Function
ErrorGrabaDepRetBancosVarios:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsgErr
End Function
'reportes -------------------------------------------------------------------
'cambio x store corresponsalia
Public Function DevolverDatos_Cabecera_108121_Corresponsalia(ByVal id As Double, Optional x As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacCabecera " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_Cabecera_108121_Corresponsalia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'-----------------------------------------------------------------------------
'--- CAMPANHAS - CASA COMERCIAL
'-----------------------------------------------------------------------------
Public Sub Insertar_Datos_CasaComercial(ByVal nombre As String, ByVal empresa As String, ByVal descripcion As String, ByVal dfinicio As Date, ByVal dfin As Date)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_CasaComercial
   
   sSql = "Exec stp_ins_CasaComercial '" & nombre & "','" & empresa & "','" & descripcion & "','" & Format(dfinicio, "yyyymmdd") & "','" & Format(dfin, "yyyymmdd") & "'"
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertar_Datos_CasaComercial:
    Err.Raise Err.Number, "Insertar Datos Casa Comercial", Err.Description
End Sub

Public Sub Actualizar_Datos_CasaComercial(ByVal id As Integer, ByVal nombre As String, ByVal empresa As String, ByVal descripcion As String, ByVal dfinicio As Date, ByVal dfin As Date)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_Datos_CasaComercial

   sSql = "Exec stp_upd_CasaComercial " & id & ",'" & descripcion & "','" & Format(dfinicio, "yyyymmdd") & "','" & Format(dfin, "yyyymmdd") & "'"
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorActualizar_Datos_CasaComercial:
    Err.Raise Err.Number, "Insertar Datos Casa Comercial", Err.Description
End Sub

Public Function DevolverDatos_CasaComercial(Optional id As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_CasaComercial " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_CasaComercial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DevolverDatos_ValidaCasaComercial(ByVal nombre As String, ByVal institucion As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_ValidaCasaComercial '" & nombre & "','" & institucion & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_ValidaCasaComercial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub ObtenerDatosCasaComercial(ByRef poDatos As ADODB.Recordset, Optional val As Integer = 0)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    If val <> 0 Then
        sSql = " exec stp_sel_CasaComercial " & val & " "
    Else
        sSql = " exec stp_sel_CasaComercial 0 "
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set poDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorObtenerDatosGrupoEconomicos:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Sub

Public Sub Insertar_Datos_ConvenioCasaComercial(ByVal id_casacom As Integer, ByVal id_camp As Integer, ByVal nContador As Integer)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
      
     sSql = "Exec stp_ins_ConvenioCasaComercial " & id_casacom & ", " & id_camp & "," & nContador
     
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
End Sub

Public Sub Eliminar_Datos_CasaComercial(ByVal id As Integer)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
   
     sSql = "Exec stp_del_CasaComercial " & id & " "
     
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
End Sub

Public Sub Eliminar_Datos_ConvenioCasaComercial(ByVal id As Integer)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
   
     sSql = "Exec stp_del_ConvenioCasaComercial " & id & " "
     
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
End Sub

Public Function DevolverDatos_ConvenioCasaComercial(ByVal id As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_ConvenioCasaComercial " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_ConvenioCasaComercial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'Solicitud de Credito
Public Function DevolverDatos_CasaComercialID(ByVal id As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_CasaComercial_Id " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatos_CasaComercialID = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'casa comercial
Public Function DevolverDatos_NumeroCasaComercial() As Integer
Dim sSql As String
Dim R As ADODB.Recordset

Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_ContarConvenioCasaComercial"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    
    If R.BOF And R.EOF Then
        DevolverDatos_NumeroCasaComercial = 0
    Else
        DevolverDatos_NumeroCasaComercial = R!Numero
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function

'*** PEAC 20100722
Public Function RecuperaCuentasParaAdmCred(ByVal psPersCod As String, ByVal pnTipoReporte As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaCuentasParaAdmCred
    
    sSql = " EXEC stp_sel_ObtieneDatosParaControlCreditosAdmCred '" & psPersCod & "'," & pnTipoReporte
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuentasParaAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuentasParaAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20100722
Public Function CargaDatosControlCreditosAdmCred(ByVal psCtaCod As String, _
                                        ByRef prsCred As ADODB.Recordset, _
                                        ByRef prsComun As ADODB.Recordset, _
                                        ByRef psEstado As String, Optional ByRef pnEstado As Long = 0, _
                                        Optional ByRef psCampania As String = "") As Boolean

Dim nEstado As Integer
Dim oCons As COMDConstantes.DCOMConstantes

On Error GoTo ErrorCargaDatosControlCreditosAdmCred

    CargaDatosControlCreditosAdmCred = True
    
    Set prsCred = RecuperaProductoAdmCred(psCtaCod)
    
    If Not (prsCred.BOF And prsCred.EOF) Then
        psEstado = prsCred!cEstado
        nEstado = prsCred!nPrdEstado
        psCampania = prsCred!cCampania
        Set prsComun = RecuperaDatosComunesAdmCred(psCtaCod, Array(nEstado, 0))
        pnEstado = nEstado
    Else
        CargaDatosControlCreditosAdmCred = False
    End If
    
    Exit Function

ErrorCargaDatosControlCreditosAdmCred:
    CargaDatosControlCreditosAdmCred = False
    Err.Raise Err.Number, "Error Carga Datos Control de Créditos", Err.Description
End Function

'*** PEAC 20100722
'MARG 20160719*******************************
Public Function RecuperaProductoAdmCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaProductoAdmCred
    sSql = " EXEC stp_sel_RecuperaProductoAdmCred2 '" & psCtaCod & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductoAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProductoAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100722
Public Function RecuperaDatosComunesAdmCred(ByVal psCtaCod As String, Optional ByVal MatEst As Variant = Nothing) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sEstados As String  ''20090212

    On Error GoTo ErrorRecuperaDatosComunesAdmCred
        
        If IsArray(MatEst) Then
            For i = 0 To UBound(MatEst)
                If i = 0 Then
                    sEstados = "" & MatEst(i) & ""
                Else
                    sEstados = sEstados & "," & MatEst(i) & ""
                End If
            Next i
        Else
            sEstados = " "
        End If
    
    sSql = " exec Stp_Sel_RecuperaDatosComunesAdmCred '" & psCtaCod & "','" & sEstados & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosComunesAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaDatosComunesAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100722
'madm 20110308 - parametros pnDesExoneraCAR
Public Sub RegistraControlAdmCred(ByVal psCtaCod As String, ByVal psFecRevision As String, Optional ByVal psObservacion As String = "", Optional ByVal pnCodExone As Integer = 0, Optional ByVal psQuienExonera As String = "", Optional ByVal psUserExonera As String = "", Optional ByVal psMovNro As String = "", Optional ByVal pnCompraDeuda As Integer = 0, Optional ByVal pRSObs As ADODB.Recordset = Nothing, Optional ByVal pRSExo As ADODB.Recordset = Nothing, Optional ByVal pnClientePref As Integer = 0, Optional ByVal pnConstitucion As Integer = 0, Optional ByVal pCompraDes As String = "", Optional ByVal pRSAuto As ADODB.Recordset = Nothing, Optional ByVal nNumAuto As Integer = 0, Optional ByVal pRSpost As ADODB.Recordset = Nothing, Optional ByVal pbPostDesembolso As Boolean = False, _
                                  Optional ByVal pbRevisaDsemb As Boolean = False)
                                  'JUEZ 20140725 Se agregó pbRevisaDsemb
Dim sSql As String
Dim sSql1 As String, sSql2 As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

'*** PEAC 20101209 - SE AGREGO EL PARAMETRO pnCompraDeuda

oCon.AbreConexion
oCon.BeginTrans

    If Not pbPostDesembolso Then 'WIOR 20120508
        'sSql = "exec stp_ins_ControlCreditosAdmCred '" & psCtaCod & "','" & psFecRevision & "','" & psObservacion & "'," & pnCodExone & ",'" & psQuienExonera & "','" & psUserExonera & "','" & psMovNro & "'," & pnCompraDeuda & "," & pnClientePref & "," & pnConstitucion & ",'" & pCompraDes & "'"
        sSql = "exec stp_ins_ControlCreditosAdmCred '" & psCtaCod & "','" & psFecRevision & "','" & psObservacion & "'," & pnCodExone & ",'" & psQuienExonera & "','" & psUserExonera & "','" & psMovNro & "'," & pnCompraDeuda & "," & pnClientePref & "," & pnConstitucion & ",'" & pCompraDes & "'," & IIf(pbRevisaDsemb, 1, 0) 'JUEZ 20140725
        oCon.Ejecutar (sSql)
    End If
    

    If Not pRSObs Is Nothing Then
        Do While Not pRSObs.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 1 & ",'" & pRSObs!cCtaCod & "','" & pRSObs!observacion & "'"
            oCon.Ejecutar (sSql)
            pRSObs.MoveNext
        Loop
    End If

    If Not pRSExo Is Nothing Then
        Do While Not pRSExo.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 2 & ",'" & pRSExo!cCtaCod & "','---'," & pRSExo!nCodExonera & ",'" & pRSExo!cCodQuienExo & "','" & pRSExo!Apoder1 & "','" & pRSExo!Apoder2 & "','" & pRSExo!Apoder3 & "'," & pRSExo!nDesExoneraCAR & ",'" & pRSExo!cDesExoneraOtros & "'"
            oCon.Ejecutar (sSql)
            pRSExo.MoveNext
        Loop
    End If

    If Not pRSAuto Is Nothing Then
        Do While Not pRSAuto.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 3 & ",'" & pRSAuto!cCtaCod & "','---'," & pRSAuto!nCodAutoriza & ",'" & pRSAuto!cCodQuienAuto & "','" & pRSAuto!Apoder1 & "','" & pRSAuto!Apoder2 & "','" & pRSAuto!Apoder3 & "'," & 4 & " ,'" & pRSAuto!AutoOtro & "'"
            oCon.Ejecutar (sSql)
            pRSAuto.MoveNext
        Loop
    End If
    
    If nNumAuto > 0 Then
           If Not pRSpost Is Nothing Then
                Do While Not pRSpost.EOF
                    sSql = " exec stp_ins_ControlAdmCredPost '" & psCtaCod & "','" & pRSpost!Agencia & "','" & pRSpost!user & "','" & pRSpost!observacion & "'"
                    oCon.Ejecutar (sSql)
                    pRSpost.MoveNext
                Loop
            End If
    End If
    
oCon.CommitTrans
oCon.CierraConexion


Set oCon = Nothing
End Sub

'MADM 20110322
Public Sub RegistraControlAdmCredMant(ByVal psCtaCod As String, Optional ByVal psMovNro As String, Optional ByVal pRSObs As ADODB.Recordset = Nothing, Optional ByVal pRSExo As ADODB.Recordset = Nothing, Optional ByVal pRSAut As ADODB.Recordset = Nothing)
Dim sSql As String
Dim sSql1 As String, sSql2 As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

sSql = " exec stp_upd_ControlCreditosAdmCred '" & psCtaCod & "','" & psMovNro & "'"
oCon.AbreConexion
oCon.BeginTrans
    oCon.Ejecutar (sSql)

    If Not pRSObs Is Nothing Then
        sSql = "exec stp_del_SelObsControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        
        Do While Not pRSObs.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 1 & ",'" & pRSObs!cCtaCod & "','" & pRSObs!observacion & "'"
            oCon.Ejecutar (sSql)
            pRSObs.MoveNext
        Loop
        'JUEZ 20140725 **********************
        sSql = " exec stp_upd_ControlCreditosAdmCredRevisaDesemb '" & psCtaCod & "',0"
        oCon.Ejecutar (sSql)
        'END JUEZ ***************************
    End If

    If Not pRSExo Is Nothing Then
        sSql = "exec stp_del_SelExoControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        
        Do While Not pRSExo.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 2 & ",'" & psCtaCod & "','---'," & pRSExo!nCodExonera & ",'" & pRSExo!cCodQuienExo & "','" & pRSExo!Apoder1 & "','" & pRSExo!Apoder2 & "','" & pRSExo!Apoder3 & "'," & pRSExo!nTipoExoneraCAR & ",'" & pRSExo!cDesExoneraOtros & "'"
            oCon.Ejecutar (sSql)
            pRSExo.MoveNext
        Loop
    Else 'RECO20141226
        sSql = "exec stp_del_SelExoControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
    End If
    If Not pRSAut Is Nothing Then
        sSql = "exec stp_del_SelAutControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        
        Do While Not pRSAut.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 3 & ",'" & psCtaCod & "','---'," & pRSAut!nCodAutoriza & ",'" & pRSAut!cCodQuienAuto & "','" & pRSAut!Apoder1 & "','" & pRSAut!Apoder2 & "','" & pRSAut!Apoder3 & "'," & 4 & ",'" & pRSAut!AutoOtro & "'"
            oCon.Ejecutar (sSql)
            pRSAut.MoveNext
        Loop
    Else 'RECO20141226
        sSql = "exec stp_del_SelAutControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
    End If
oCon.CommitTrans
oCon.CierraConexion
Set oCon = Nothing
End Sub
'MADM 20110309
Public Function BuscaExoAdmCred(ByVal psCtaCod As String) As Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaExoAdmCred
    
    sSql = " EXEC stp_sel_BuscaExoControlAdmCred '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BuscaExoAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorBuscaExoAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END MADM

'*** PEAC 20100722
Public Function BuscaRegistroControlAdmCred(ByVal psCtaCod As String, Optional ByRef pnTpoRevision As Integer = 0) As Boolean
'BuscaRegistroVisitaClienteAdmCred
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaRegistroControlAdmCred
    
    sSql = " EXEC stp_sel_BuscaRegistroControlAdmCred '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        BuscaRegistroControlAdmCred = False
        pnTpoRevision = 0
    Else
        pnTpoRevision = IIf(IsNull(rs!nTpoRevision) = True, 0, rs!nTpoRevision)
        BuscaRegistroControlAdmCred = True
    End If
    Exit Function
ErrorBuscaRegistroControlAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneQuienExoneraAdmCred() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneQuienExoneraAdmCred
    
    sSql = " EXEC stp_sel_ObtieneQuienExoneraAdmCred "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneQuienExoneraAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneQuienExoneraAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneUserExoneraAdmCred(ByVal psCargoCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneUserExoneraAdmCred
    
    If Len(psCargoCod) = 0 Then
        sSql = " EXEC stp_sel_ObtieneUserExonera "
    Else
        sSql = " EXEC stp_sel_ObtieneUserExonera '" & psCargoCod & "'"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneUserExoneraAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneUserExoneraAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100722
Public Sub RegistraVisitaClienteAdmCred(ByVal psCtaCod As String, ByVal psFecRevision As String, Optional ByVal psDireccion As String = "", Optional ByVal psEntrevistado As String = "", Optional ByVal pnCodVerifi As Integer = 0, Optional ByVal psVerifGar As String = "", Optional ByVal pnOpiCaja As Integer = 0, Optional ByVal pnOpiAna As Integer = 0, Optional ByVal psComentarios As String = "", Optional ByVal psConclusion As String = "", Optional ByVal psMovNro As String = "")
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_ins_VisitaClienteAdmCred '" & psCtaCod & "','" & psFecRevision & "','" & psDireccion & "','" & psEntrevistado & "'," & pnCodVerifi & ",'" & psVerifGar & "'," & pnOpiAna & "," & pnOpiCaja & ",'" & psComentarios & "','" & psConclusion & "','" & psMovNro & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub

'*** PEAC 20100722
Public Function BuscaRegistroVisitaClienteAdmCred(ByVal psCtaCod As String, ByVal psFecha As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaRegistroVisitaClienteAdmCred
    
    sSql = " EXEC stp_sel_BuscaRegistroVisitaClienteAdmCred '" & psCtaCod & "','" & Format(psFecha, "yyyymmdd") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        BuscaRegistroVisitaClienteAdmCred = False
    Else
        BuscaRegistroVisitaClienteAdmCred = True
    End If
    
    Exit Function

ErrorBuscaRegistroVisitaClienteAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'1*** MADM 20110802 - PEAC 20100723
Public Function RecuperaDatosAprobacionCreditos(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pnCliPref As Integer = 0, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer
    
    On Error GoTo ErrorRecuperaDatosAprobacionCreditos
    
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    'sSql = " exec stp_sel_ReporteControlCreditos '" & psAge & "','" & psDel & "','" & psAl & "'," & pnCliPref
    'MADM 20110802
    sSql = " exec stp_sel_ReporteControlCreditos '" & psAge & "','" & psDel & "','" & psAl & "'," & pnCliPref & ",'" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosAprobacionCreditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosAprobacionCreditos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneAgenciasAdmCred() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneAgenciasAdmCred
    
    sSql = " EXEC stp_sel_ObtieneAgenciasAdmCred "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneAgenciasAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneAgenciasAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneCredDesembolsadosPorDia(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosAprobacionCreditos
    
    sSql = " exec stp_sel_ObtieneCredDesembolsadosPorDia '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCredDesembolsadosPorDia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosAprobacionCreditos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneCredDesembolsadosConObservaciones(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer
    
    On Error GoTo ErrorObtieneCredDesembolsadosConObservaciones
    
   sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    
    sSql = " exec stp_sel_ObtieneCredDesembolsadosConObservaciones '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCredDesembolsadosConObservaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredDesembolsadosConObservaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100723
'MADM 20110802 (pMatProd) - ObtieneCreditosCompraDeuda
Public Function ObtieneCredDesembolsadosConExoneraciones(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer

    On Error GoTo ErrorObtieneCredDesembolsadosConExoneraciones
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
      
    sSql = " exec stp_sel_ObtieneCredDesembolsadosConExoneraciones '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneCredDesembolsadosConExoneraciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredDesembolsadosConExoneraciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MADM 20110802 *** PEAC 20100723
Public Function ObtieneCantidadObservacionesPorAnalista(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer
    
    On Error GoTo ErrorObtieneCantidadObservacionesPorAnalista
        
   sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sSql = " exec stp_sel_ObtieneObservacionesPorAnalista '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCantidadObservacionesPorAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCantidadObservacionesPorAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MADM 20110803 *** PEAC 20100723
Public Function ImprimeCantidadExoneracionesPorAnalista(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer
    
    On Error GoTo ErrorImprimeCantidadExoneracionesPorAnalista
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sSql = " exec stp_sel_ObtieneExoneracionesPorAnalista '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ImprimeCantidadExoneracionesPorAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorImprimeCantidadExoneracionesPorAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MADM 20110802 *** PEAC 20100723
Public Function ImprimeCantidadExoneracionesPorRespApr(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer
    
    On Error GoTo ErrorImprimeCantidadExoneracionesPorRespApr
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sSql = " exec stp_sel_ObtieneExoneracionesPorRespAprob '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ImprimeCantidadExoneracionesPorRespApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorImprimeCantidadExoneracionesPorRespApr:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20100722
Public Function CargaDatosVisitasClientesAdmCred(ByVal psCtaCod As String, _
                                        ByRef prsCred As ADODB.Recordset, _
                                        ByRef prsComun As ADODB.Recordset, _
                                        ByRef psEstado As String, Optional ByRef pnEstado As Long = 0) As Boolean

Dim nEstado As Integer
Dim oCons As COMDConstantes.DCOMConstantes

On Error GoTo ErrorCargaDatosVisitasClientesAdmCred

    CargaDatosVisitasClientesAdmCred = True
    
    Set prsCred = RecuperaProductoAdmCred(psCtaCod)
    
    If Not (prsCred.BOF And prsCred.EOF) Then
        psEstado = prsCred!cEstado
        nEstado = prsCred!nPrdEstado
        Set prsComun = RecuperaDatosComunesAdmCred(psCtaCod, Array(nEstado, 0))
        pnEstado = nEstado
    Else
        CargaDatosVisitasClientesAdmCred = False
    End If
    
    Exit Function

ErrorCargaDatosVisitasClientesAdmCred:
    CargaDatosVisitasClientesAdmCred = False
    Err.Raise Err.Number, "Error Carga Datos Control de Visitas", Err.Description
End Function

'*** PEAC 20100723
Public Function ObtieneVisitasClientes(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneVisitasClientes
    
    sSql = " exec stp_sel_ReporteVisitasClientesAdmCred '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneVisitasClientes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneVisitasClientes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20101229
Public Function BuscaObsAdmCred(ByVal psCtaCod As String) As Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaObsAdmCred
    
    sSql = " EXEC stp_sel_BuscaObsControlAdmCred '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BuscaObsAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20101229
Public Sub RegularizaObsAdmCred(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal pRSObs As ADODB.Recordset, ByVal pbRevisaDesemb As Boolean)

'vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew

Dim sSql As String
Dim sSql1 As String, sSql2 As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

sSql = " exec stp_upd_ControlCreditosAdmCred '" & psCtaCod & "','" & psMovNro & "'"

oCon.AbreConexion
'oCon.BeginTrans

    oCon.Ejecutar (sSql)
    'JUEZ 20140725 **********************
    sSql = " exec stp_upd_ControlCreditosAdmCredRevisaDesemb '" & psCtaCod & "'," & IIf(pbRevisaDesemb, 1, 0)
    oCon.Ejecutar (sSql)
    'END JUEZ ***************************
    If Not pRSObs Is Nothing Then
        Do While Not pRSObs.EOF
            sSql = " exec stp_upd_ControlAdmCredObsExo '" & pRSObs!cCtaCod & "','" & pRSObs!observacion & "'," & pRSObs!OK
            oCon.Ejecutar (sSql)
            pRSObs.MoveNext
        Loop
    End If


'oCon.CommitTrans

oCon.CierraConexion

Set oCon = Nothing
End Sub

'*** PEAC 20101229
Public Function RecuperaCuentasObsParaAdmCred(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaCuentasObsParaAdmCred
    
    sSql = " EXEC RecuperaCuentasObsParaAdmCred '" & psPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuentasObsParaAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuentasObsParaAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'MADM 20110308
Public Function RecuperaCuentasExoObsParaAdmCred(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaCuentasObsParaAdmCred
    
    sSql = " EXEC RecuperaCuentasExoParaAdmCred '" & psPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuentasExoObsParaAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuentasObsParaAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtieneCreditosConstitucion(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneCreditosConstitucion
    
    sSql = " exec stp_sel_ObtieneCreditosConstitucion '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCreditosConstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCreditosConstitucion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'MADM 20111109
Public Function RecuperaPagosEnLoteBN(ByVal tipo As Integer, ByVal pdFecSist As Date, ByVal pdFecOper As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    If tipo = 0 Then
        sSql = "exec stp_sel_LogCobrosExtorno '" & Format(pdFecSist, "YYYYMMDD") & "', '" & Format(pdFecOper, "YYYYMMDD") & "' "
    ElseIf tipo = 1 Then
        sSql = "exec stp_sel_LogCorresponsaliaExtorno '" & Format(pdFecSist, "YYYYMMDD") & "', '" & Format(pdFecOper, "YYYYMMDD") & "' "
    Else
        sSql = "exec stp_sel_LogBCPExtorno '" & Format(pdFecSist, "YYYYMMDD") & "', '" & Format(pdFecOper, "YYYYMMDD") & "' "
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagosEnLoteBN = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
Public Function RecuperaCreditosBNCorrexConv(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    'sSql = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro & " and NOT cOpecod LIKE '99%'"
    sSql = "exec stp_sel_DatosCredxMovConxCorresponsalia " & pnMovNro & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosBNCorrexConv = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'END MADM

'*** PEAC 20101229
Public Function ObtieneCreditosCompraDeuda(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneCreditosCompraDeuda
    
    sSql = " exec stp_sel_ObtieneCreditosCompraDeuda '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCreditosCompraDeuda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCreditosCompraDeuda:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MADM 20120329
Public Function ObtieneCredDesembolsadosConAutorizaciones(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer

    On Error GoTo ErrorObtieneCredDesembolsadosConExoneraciones
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
      
    sSql = " exec stp_sel_ObtieneCredDesembolsadosConAutorizaciones '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneCredDesembolsadosConAutorizaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredDesembolsadosConExoneraciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtieneCredDesembolsadosConPostDesembolso(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer

    On Error GoTo ErrorObtieneCredDesembolsadosConExoneraciones
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
      
    sSql = " exec stp_sel_ObtieneCredDesembolsadosConPostD '" & psAge & "','" & psDel & "','" & psAl & "','" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneCredDesembolsadosConPostDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredDesembolsadosConExoneraciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtieneCredCFObservaciones(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer

    On Error GoTo ErrorObtieneCredCFObservaciones
        
    sSql = " exec stp_sel_ObtieneCartasFianzasObservaciones '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneCredCFObservaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredCFObservaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaAdmCredRefinan(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaAdmCredRefinan
    
    sSql = " EXEC stp_sel_ObtieneAdmCredRefinan '" & pcCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAdmCredRefinan = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAdmCredRefinan:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END MADM
'WIOR 20120620 *******************************************
Public Function ObtieneCredCFPostDesembolso(ByVal psCtaCod As String) As Boolean

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As Recordset

    On Error GoTo ErrorObtieneCredCFPostDesembolso
        
    sSql = " select nPostNro ,cCtaCod ,cAgeCod ,cUsu,CDescripcionOtro " & _
            " from ControlCreditosPostAdmCred where cCtaCod ='" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set rs = oConecta.CargaRecordSet(sSql)
    
     If (rs.EOF And rs.BOF) Then
        ObtieneCredCFPostDesembolso = False
    Else
        ObtieneCredCFPostDesembolso = True
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredCFPostDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ************************************************

'*************RECO 20131012 ERS055-2013*******************
Public Function dObtieneCreditosDePersona(ByVal psPersCod As String, _
         Optional ByVal psAgencia As String = "") As ADODB.Recordset
'Obtiene Listado Creditos en Recuperaciones  de una Persona
Dim lrs As ADODB.Recordset
Dim sSql As String
Dim lsAgencia As String
Dim oConecta As COMConecta.DCOMConecta

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError

sSql = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cRelacion, isnull(CN4.nConsValor,0) nTipoProdCod, ISNULL(CN4.cConsDescripcion,'') as cTipoProdDescrip, " _
        & "UPPER(T3.cConsDescripcion) cMoneda " _
        & "FROM ProductoPersona PP INNER JOIN Producto P  " _
        & "INNER JOIN Colocaciones CO ON P.cCtaCod = CO.cCtaCod " _
        & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
        & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN Constante CN4 " _
        & "ON convert(int,substring(cTpoProdCod,1,2)+'0') = CN4.nConsValor AND CN4.nConsCod = " & gTpoProducto & " INNER JOIN " & csCentralCom & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
        & "T3.nConsCod = " & gMoneda _
        & " AND P.cCtaCod like '___" & lsAgencia & "%' " _
        & "AND P.nPrdEstado in ( 2020, 2021, 2022, 2030, 2031, 2032) "
sSql = sSql & " ORDER BY PP.cCtaCod"

 Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

Set lrs = oConecta.CargaRecordSet(sSql)

Set dObtieneCreditosDePersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function
Public Function RecuperaValorEdificacion(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo ErrorBuscaObsAdmCred



sSql = "EXEC stp_sel_RecuperaValorEdificacion '" & psNumGarant & "'"

 Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaValorEdificacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerUltimaCuotaPag(ByVal psCuenta As String) As ADODB.Recordset
    Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo ErrorBuscaObsAdmCred



sSql = "EXEC stp_sel_ObtenerUltimaCuotaPag '" & psCuenta & "'"

 Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerUltimaCuotaPag = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ObtenerCotizacionSeguroXAgencia(ByVal psAgeCod As String, ByVal pnInmueble As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo ErrorBuscaObsAdmCred



sSql = "EXEC stp_sel_ObtenerCotizacionSeguro '" & psAgeCod & "'," & pnInmueble

 Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerCotizacionSeguroXAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'Public Sub MantenimientoGastoPolizaContraincendio(ByVal lsCtaCod As String, ByVal cUser As String, ByVal cAgeCod As String _
'                , ByVal dFechaReg As String, ByVal nCuotaDesde As Integer, ByVal nCuotaHasta As Integer, ByVal nPrimaNeta As Double _
'                , ByVal nPrimaTotal As Double, ByVal nPrimaTotalCT As Double, ByVal nMontoPrimaCuota As Double)
'
'    Dim sSql As String
'    Dim oConecta As COMConecta.DCOMConecta
'    Dim rs As Recordset
'
'    On Error GoTo ErrorBuscaObsAdmCred
'
'
'
'    sSql = "EXEC stp_ins_MantenimientoGastoPolizaContraincendio '" & psAgeCod & "'," & pnInmueble
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    oConecta.CargaRecordSet (sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'    Exit Function
'
'ErrorBuscaObsAdmCred:
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'
'End Sub
'*********************END RECO****************************

'*******************RECO 20131028 ERS133
Public Function BuscaObsPosAdmCred(ByVal psCtaCod As String) As Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaObsAdmCred
    
    sSql = " EXEC stp_sel_BuscaObsPosControlAdmCred '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BuscaObsPosAdmCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'*********************END RECO****************

'************RECO 20131029*************************************
Public Sub RegularizaObsPosAdmCred(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal pRSObs As ADODB.Recordset)

'vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew

Dim sSql As String
Dim sSql1 As String, sSql2 As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

sSql = " exec stp_upd_ControlCreditosAdmCred '" & psCtaCod & "','" & psMovNro & "'"

oCon.AbreConexion
'oCon.BeginTrans

    oCon.Ejecutar (sSql)

    If Not pRSObs Is Nothing Then
        Do While Not pRSObs.EOF
            sSql = " exec stp_upd_ControlAdmCredObsPosExo '" & pRSObs!cCtaCod & "','" & pRSObs!observacion & "'," & pRSObs!OK
            oCon.Ejecutar (sSql)
            pRSObs.MoveNext
        Loop
    End If


'oCon.CommitTrans

oCon.CierraConexion

Set oCon = Nothing
End Sub
'****************END RECO**************************************
'********************RECO 20131029 ERS141*****************
Public Function ObtieneDatosCredEcotaxi(ByVal spCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim lsAgencia As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_sel_ObtieneDatosCredEcotaxi '" & spCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

    Set lrs = oConecta.CargaRecordSet(sSql)

    Set ObtieneDatosCredEcotaxi = lrs
    Set lrs = Nothing

    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function
    

'*********************END RECO****************************

'WIOR 20140205 ****************************************
Public Function ObtieneAsignacionesSaldo(ByVal psAge As String, ByVal psDel As String, ByVal psAl As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sCadProd As String
    Dim i As Integer

    On Error GoTo ErrorObtieneCredCFObservaciones
        
    sSql = " exec stp_sel_ReporteAsignarSaldos '" & psAge & "','" & psDel & "','" & psAl & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneAsignacionesSaldo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredCFObservaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN **********************************************
'********************RECO 20131106 ERS150*****************
Public Function ObtieneDatosCredPersona(ByVal spCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim lsAgencia As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_sel_ObtieneDatosCredPersona '" & spCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

    Set lrs = oConecta.CargaRecordSet(sSql)

    Set ObtieneDatosCredPersona = lrs
    Set lrs = Nothing

    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<ObtieneDatosCredPersona>>", Err.Description
End Function

Public Sub RegistraConvenioCredPosDesembolso(ByVal psCtaCod As String, ByVal psPersCodInst As String, ByVal psCodModular As String, ByVal psCargo As String, ByVal psCARBEN As String, ByVal psT_Plani As String, ByVal psFecReg As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_ins_RegistraConvenioCredPosDesembolso '" & psCtaCod & "','" & psPersCodInst & "','" & psCodModular & " ','" & psCargo & "','" & psCARBEN & "','" & psT_Plani & "','" & psFecReg & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    Exit Sub

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<RegistraConvenioCredPosDesembolso>>", Err.Description
End Sub


Public Function ObtieneDatosCreditoConvenio(ByVal spCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim lsAgencia As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_sel_ObtieneDatosCreditoConvenio '" & spCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

    Set lrs = oConecta.CargaRecordSet(sSql)

    Set ObtieneDatosCreditoConvenio = lrs
    Set lrs = Nothing

    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<ObtieneDatosCreditoConvenio>>", Err.Description
End Function

Public Function ObtieneCtaConvenioPostDesembolso(ByVal spCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim lsAgencia As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_sel_ObtieneCtaConvenioPostDesembolso '" & spCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

    Set lrs = oConecta.CargaRecordSet(sSql)

    Set ObtieneCtaConvenioPostDesembolso = lrs
    Set lrs = Nothing

    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Convenio Pos Desembolso <<ObtieneCtaConvenioPostDesembolso>>", Err.Description
End Function

Public Sub EliminaConvenioCredPosDesembolso(ByVal psCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_ins_EliminaConvenioCredPosDesembolso '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    Exit Sub

dError:
    Err.Raise Err.Number, "Elimina convenio pos desembolso <<EliminaConvenioCredPosDesembolso>>", Err.Description
End Sub
'*********************END RECO****************************
'********************RECO20131112 ERS154******************
Public Sub RegistrarHojaRuta(psFecha, pnNumVisita, psPersCodCliente, psHora, pnEstado, psUserAnalista, psUserRegistra)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_ins_RegistrarHojaRuta '" & Format(psFecha, "yyyy/MM/dd") & "'," & pnNumVisita & ",'" & psPersCodCliente & "','" & psHora & "'," & pnEstado & ",'" & psUserAnalista & "','" & psUserRegistra & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    Exit Sub

dError:
    Err.Raise Err.Number, "Registra Datos Hoja de Ruta en <<RegistrarHojaRuta>>", Err.Description
End Sub

Public Function ObtenerDatosHojaRutaXAnalista(ByVal spUserAnalista As String, ByVal spFecha As String, ByVal npOpeTpo As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim lsAgencia As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo dError
    
    sSql = "stp_sel_ObtenerDatosHojaRutaXAnalista '" & spUserAnalista & "','" & Format(spFecha, "yyyy/MM/dd") & "'," & npOpeTpo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion

    Set lrs = oConecta.CargaRecordSet(sSql)

    Set ObtenerDatosHojaRutaXAnalista = lrs
    Set lrs = Nothing

    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Hoja de Ruta en <<ObtenerDatosHojaRutaXAnalista>>", Err.Description
End Function

Public Function ListarAnalistasXAge(ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
        
    On Error GoTo dError
    
    sSql = "stp_sel_ListarAnalistasXAgencia '" & psCodAge & "'"
     oConecta.AbreConexion
    Set ListarAnalistasXAge = oConecta.CargaRecordSet(sSql)
      
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Lista Analista en <<ListarAnalistasXAge>>", Err.Description
End Function

Public Function ObtenerDatosPersonaXUser(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
        
    On Error GoTo dError
    
    sSql = "stp_sel_ObtenerDatosPersonaXUser '" & psUser & "'"
     oConecta.AbreConexion
    Set ObtenerDatosPersonaXUser = oConecta.CargaRecordSet(sSql)
      
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Lista Analista en <<ListarAnalistasXAge>>", Err.Description
End Function

Public Sub ActualizarVisitaAnalista(ByVal pnVisita As Integer, ByVal pnResultado As Integer, ByVal psObservacion As String, ByVal pnHojaRutaCod As Long)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
        
    On Error GoTo dError
    
    'sSql = "stp_upd_ActualizarHojaRutaAnalista '" & psUser & "','" & psFecha & "'," & pnVisita & "," & pnResultado & ",'" & psObservacion & "','" & psFecResultado & "'"
    sSql = "stp_upd_ActualizarHojaRutaAnalista " & pnVisita & "," & pnResultado & ",'" & psObservacion & "'," & pnHojaRutaCod
     oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
      
    Exit Sub
dError:
    Err.Raise Err.Number, "Actualiza Datos Hoja de Ruta por Analista en <<ActualizarVisitaAnalista>>", Err.Description
End Sub

Public Function ObtieneNumeroRutaEnDia(ByVal psFecha As String, ByVal psUserCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
        
    On Error GoTo dError
    
    sSql = "STP_SEL_ObtieneNumeroRutaEnDia '" & psFecha & "','" & psUserCod & "'"
     oConecta.AbreConexion
    Set ObtieneNumeroRutaEnDia = oConecta.CargaRecordSet(sSql)
      
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Numero Ruta En Dia en<<ObtieneNumeroRutaEnDia>>", Err.Description
End Function

Public Function ObtieneNumeroRutaDiaAnt(ByVal psFecha As String, ByVal psUserCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
        
    On Error GoTo dError
    
    sSql = "STP_SEL_ObtieneNumeroRutaDiaAnt '" & psFecha & "','" & psUserCod & "'"
     oConecta.AbreConexion
    Set ObtieneNumeroRutaDiaAnt = oConecta.CargaRecordSet(sSql)
      
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Numero Ruta Dia Ant en<<ObtieneNumeroRutaDiaAnt>>", Err.Description
End Function
'********************END RECO*****************************
'FRHU 20140212 RQ14010 ****************************************
Public Function ObtieneAgenciasActivas() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneAgenciasActivas
        
    sSql = " exec stp_sel_AgenciasActivas "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneAgenciasActivas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneAgenciasActivas:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ObtieneJefeTerritorial() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneJefeTerritorial
        
    sSql = " exec stp_sel_JefeNegociosTerritoriales "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ObtieneJefeTerritorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneJefeTerritorial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ValidarJefeNegocio(ByVal pcAgeCod As String, ByVal pcPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorValidarJefeNegocio
        
    sSql = " exec stp_sel_ValidarAgenciaJefeTerritorial '" & pcAgeCod & "','" & pcPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ValidarJefeNegocio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorValidarJefeNegocio:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ModificaAgenciaJNTerritorial(ByVal pcAgeCod As String, ByVal pcPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorModificaAgenciaJNTerritorial
        
    sSql = " exec stp_upd_AgenciaJefeTerritorial '" & pcAgeCod & "','" & pcPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set ModificaAgenciaJNTerritorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorModificaAgenciaJNTerritorial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'FRHU 20140212 **********************************************
'FRHU 20140221 RQ14016
Public Function DatosPosicionClienteDeudaEnEntidades(ByVal pNumDocID As String, ByVal pnTipoDocID As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDatosPosicionClienteDeudaEnEntidades
    sSql = "exec stp_sel_DatosPosicionClienteDeudaEnEntidades '" & pNumDocID & "'," & pnTipoDocID
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set DatosPosicionClienteDeudaEnEntidades = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDatosPosicionClienteDeudaEnEntidades:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'FIN FRHU 20140221 RQ14016

'JUEZ 201420725 ********************************************************************
Public Function VerificaRevisaDesembControlAdmCred(ByVal psCtaCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorVerificaRevisaDesembControlAdmCred
    
    sSql = " EXEC stp_sel_BuscaRegistroControlAdmCred '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        VerificaRevisaDesembControlAdmCred = False
    Else
        VerificaRevisaDesembControlAdmCred = IIf(IsNull(rs!bRevisaDesemb), False, rs!bRevisaDesemb)
    End If
    Exit Function
ErrorVerificaRevisaDesembControlAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END JUEZ **************************************************************************


Public Function CreditoReprogramado(ByVal psCtaCod As String) As Boolean

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorCreditoReprogramado

    sSql = " EXEC stp_sel_CreditoReprogramado '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    CreditoReprogramado = False
    If Not (rs.EOF And rs.BOF) Then
        If (rs!cReprogracion = "SI") Then
            CreditoReprogramado = True
        End If
    End If
    
    Exit Function
ErrorCreditoReprogramado:
    CreditoReprogramado = False
    Err.Raise Err.Number, "Error Carga Datos Control de Créditos", Err.Description
End Function
'RECO20141009*************************************************************
Public Function DevulveExoneracionesCredControlAdm(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDatosExoneraciones
    sSql = "exec stp_sel_DevulveExoneracionesCredControlAdm '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set DevulveExoneracionesCredControlAdm = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDatosExoneraciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function DevulveAutorizacionesCredControlAdm(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDatosAutorizaciones
    sSql = "exec stp_sel_DevulveAutorizacionesCredControlAdm '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set DevulveAutorizacionesCredControlAdm = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDatosAutorizaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN*****************************************************************
'RECO20150331 ERS010-2015**************************************
Public Function ObtieneCredAdmAgeCodAutirza(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneCredAdmAgeCodAutirza
    sSql = "exec stp_sel_ObtieneCredAdmAgeCodAutoriza '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set ObtieneCredAdmAgeCodAutirza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneCredAdmAgeCodAutirza:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub ControlAdmCred(ByVal psCtaCod As String, ByVal psFecRevision As String, ByVal psMovNro As String, ByVal pRSObs As ADODB.Recordset, ByVal pRSExo As ADODB.Recordset, _
                          ByVal pRSAuto As ADODB.Recordset, ByVal pbPostDesembolso As Boolean, ByVal pbRevisaDsemb As Boolean, ByVal pnTpoRevision As Integer)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
Dim oMov As New COMDCredito.DCOMCredActBD
oCon.AbreConexion
oCon.BeginTrans

    Call oMov.InsertaMov(psMovNro, IIf(pbPostDesembolso, 100924, 100923), IIf(pbPostDesembolso, "Reg. Control Post Desembolso", "Reg. Control Pre Desembolso") & psCtaCod, gMovEstContabNoContable, gMovFlagVigente)
    
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If Not pbPostDesembolso Or objProducto.GetResultadoCondicionCatalogo("N0000097", Mid(psCtaCod, 6, 3)) Then
    'If Not pbPostDesembolso Or Mid(psCtaCod, 6, 3) = "703" Then
    '**ARLO20180712 ERS042 - 2018
        Dim rs As New ADODB.Recordset
        Set rs = oCon.CargaRecordSet("stp_sel_VerificaExisteControl '" & psCtaCod & "'")
        If (rs.BOF And rs.EOF) Then
            sSql = "exec stp_ins_CredAdmControl '" & psCtaCod & "','" & psFecRevision & "','" & psMovNro & "'," & IIf(pbRevisaDsemb, 1, 0)
            oCon.Ejecutar (sSql)
        End If
        If pbRevisaDsemb = True Then
            sSql = "exec stp_upd_ActualizaTipoRevision '" & psCtaCod & "'," & pnTpoRevision
            oCon.Ejecutar (sSql)
        End If
    End If
    
    If Not pRSObs Is Nothing And pbPostDesembolso = False Then
        sSql = "exec stp_del_SelObsControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        Do While Not pRSObs.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 1 & ",'" & pRSObs!cCtaCod & "','" & pRSObs!observacion & "','','','','','','',''," & pRSObs!nRegulariza
            oCon.Ejecutar (sSql)
            pRSObs.MoveNext
        Loop
    End If

    'If Not pRSExo Is Nothing And pbPostDesembolso = False Then
    If Not pRSExo Is Nothing Then
        sSql = "exec stp_del_SelExoControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        Do While Not pRSExo.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 2 & ",'" & pRSExo!cCtaCod & "',''," & pRSExo!nCodExonera & ",'" & pRSExo!cCodQuienExo & "','" & pRSExo!Apoder1 & "','" & pRSExo!Apoder2 & "','" & pRSExo!Apoder3 & "'," & pRSExo!nDesExoneraCAR & ",'" & pRSExo!cDesExoneraOtros & "'"
            oCon.Ejecutar (sSql)
            pRSExo.MoveNext
        Loop
    Else
        If Not pbPostDesembolso Then
            sSql = "exec stp_del_SelExoControlAdmCred '" & psCtaCod & "'"
            oCon.Ejecutar (sSql)
         End If
    End If

    'If Not pRSAuto Is Nothing And pbPostDesembolso = False Then
    If Not pRSAuto Is Nothing Then
        sSql = "exec stp_del_SelAutControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
        Do While Not pRSAuto.EOF
            sSql = " exec stp_ins_ControlAdmCredObsExo " & 3 & ",'" & pRSAuto!cCtaCod & "',''," & pRSAuto!nCodAutoriza & ",'" & pRSAuto!cCodQuienAuto & "','" & pRSAuto!Apoder1 & "','" & pRSAuto!Apoder2 & "','" & pRSAuto!Apoder3 & "'," & 4 & " ,'" & pRSAuto!AutoOtro & "'"
            oCon.Ejecutar (sSql)
            pRSAuto.MoveNext
        Loop
    Else
        sSql = "exec stp_del_SelAutControlAdmCred '" & psCtaCod & "'"
        oCon.Ejecutar (sSql)
    End If
    
    If pbPostDesembolso Then
        If Not pRSObs Is Nothing Then
            sSql = "exec stp_del_ControlCreditosPostAdmCred '" & psCtaCod & "'"
            oCon.Ejecutar (sSql)
            Do While Not pRSObs.EOF
                sSql = " exec stp_ins_ControlAdmCredPost '" & psCtaCod & "','" & pRSObs!observacion & "'," & pRSObs!nRegulariza & "," & pRSObs!nAnalistaObs
                oCon.Ejecutar (sSql)
                pRSObs.MoveNext
            Loop
        End If
    End If
    
oCon.CommitTrans
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Function DevuelveObsPostDesembolso(ByVal psCtaCod As String) As Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorBuscaObsAdmCred
    
    sSql = " EXEC stp_sel_DevuelveObsPostDesembolso '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveObsPostDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorBuscaObsAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN******************************************************

'RECO20150703 ERS034-2015**************************************
Public Function GeneraTramaMicroSeguroMultiriesgoMYPE(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As Recordset
    On Error GoTo ErrorGeneraTramaMYPE
        
        sSql = " EXEC stp_sel_GeneraTramaMicroSeguroMultiriesgoMYPE " & pnMes & "," & pnAnio
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GeneraTramaMicroSeguroMultiriesgoMYPE = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
    
ErrorGeneraTramaMYPE:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RegistraTramaSegMYPE(ByVal psTpoDeclaracion As String, ByVal psFecProceso As String, ByVal psFecIni As String, ByVal psFecFin As String, ByVal psFormaPago As String, _
    ByVal pnValorTotalAseg As Integer, ByVal pnEstado As Integer) As Long
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrorRegistroTramaCabecera
        
        sSql = " EXEC spt_ins_RegistraTramaSegMYPE '" & psTpoDeclaracion & "','" & psFecProceso & "','" & psFecIni & "','" & psFecFin & "','" & psFormaPago & "'," & pnValorTotalAseg & "," & pnEstado
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        If Not (rs.EOF And rs.BOF) Then
            RegistraTramaSegMYPE = rs!nTramaID
        Else
            RegistraTramaSegMYPE = 0
        End If
        Set oConecta = Nothing
        
        Exit Function
    
ErrorRegistroTramaCabecera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub RegistraTramaSegMYPEDet(ByVal pnTramaID As Integer, ByVal psPlanCod As String, ByVal pnValorAseg As Double, ByVal pnPrimaAnual As Double, ByVal pnPrimaMes As Double, _
    psNumPoliza As String, ByVal psNumGarant As String, ByVal psPersCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As Recordset
    On Error GoTo ErrorRegistroTramaDetalle
        
        sSql = " EXEC spt_ins_RegistraTramaSegMYPEDet '" & pnTramaID & "','" & psPlanCod & "'," & pnValorAseg & "," & pnPrimaAnual & "," & pnPrimaMes & ",'" & psNumPoliza & "','" & psNumGarant & "','" & psPersCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.Ejecutar sSql
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorRegistroTramaDetalle:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
Public Function ValidaExisteTrama(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorValidaExisteTrama
        
        sSql = " EXEC stp_sel_ValidaExisteTrama " & pnMes & "," & pnAnio
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ValidaExisteTrama = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
    
ErrorValidaExisteTrama:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN *****************************************************
'EJVG20150702 ***
Public Function dObtieneCreditosDePersonaxPolizaManual(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo dError
    sSql = "EXEC stp_sel_ERS0632014_CreditoxPolizaManual '" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set dObtieneCreditosDePersonaxPolizaManual = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneCreditosDePersonaxPolizaManual>>", Err.Description
End Function
Public Function dObtieneCreditosDePersonaxPoliza(ByVal pnTpoInicio As Integer, ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo dError
    sSql = "EXEC stp_sel_ERS0632014_CreditoxPoliza " & pnTpoInicio & ",'" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set dObtieneCreditosDePersonaxPoliza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneCreditosDePersonaxPoliza>>", Err.Description
End Function
Public Function dListaSolicitudExoneraCobertura(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo dError
    sSql = "EXEC stp_sel_ERS0632014_ExoneraCobertura '" & psUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set dListaSolicitudExoneraCobertura = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error en <<dListaSolicitudExoneraTasa>>", Err.Description
End Function
'END EJVG *******

'LUCV20160613, Segun: ERS004-2016 **********
Public Function DevuelveAprobacionCreditoRefinanciado(ByVal pcCtaCod As String, psAgencia As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ValidaPagoCreditoRefinanciado '" & pcCtaCod & "','" & psAgencia & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveAprobacionCreditoRefinanciado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'Fin LUCV20160613 **********
'ARLO 20170904 ********************************************************************
Public Function VerificaClienteCampania(ByVal psCtaCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorVerificaRevisaDesembControlAdmCred
    
    sSql = " EXEC stp_sel_ERS0452017_VerificaClienteCampania '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        VerificaClienteCampania = False
    Else
        VerificaClienteCampania = True
    End If
    Exit Function
ErrorVerificaRevisaDesembControlAdmCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'ARLO 20170904 ********************************************************************
'ARLO 20170910 ********************************************************************
Public Function VerificaClienteCampaniaSolicitud(ByVal pscPerscod As String) As ADODB.Recordset 'ARLO 20170925
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorVerificaClienteCampaniaSolicitud
    
    sSql = " EXEC stp_sel_ERS0452017_VerificaClienteCampaniaSolicitud '" & pscPerscod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaClienteCampaniaSolicitud = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
'   COMENTADO POR ARLO 20170925
'    If (rs.EOF And rs.BOF) Then
'        VerificaClienteCampaniaSolicitud = False
'    Else
'        VerificaClienteCampaniaSolicitud = True
'    End If
'    Exit Function
'   COMENTADO POR ARLO 20170925
    Exit Function
ErrorVerificaClienteCampaniaSolicitud:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'ARLO 20170910 ********************************************************************
'ARLO 20171127 ********************************************************************
Public Function VerificaCampania(ByVal psIdCampana As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorVerificaCampania
    
    sSql = " EXEC stp_sel_ERS0452017_VerificaCampania '" & psIdCampana & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        VerificaCampania = False
    Else
        VerificaCampania = True
    End If
    Exit Function
ErrorVerificaCampania:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'**ARLO20171127 INICIO ERS070 - 2017
Public Function obtenerCreditosCancelados(ByVal pscPerscod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorObtenerCreditosCancelados
    
    sSql = " EXEC stp_sel_ObtenerCreditoCancelado '" & pscPerscod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set obtenerCreditosCancelados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerCreditosCancelados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function obtenerCalificacionRCC(ByVal pscPerscod As String, ByVal pscCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorObtenerCreditosCancelados
    
    sSql = " EXEC stp_sel_obtenerCalificacionRCC '" & pscPerscod & "','" & pscCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set obtenerCalificacionRCC = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerCreditosCancelados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function obtieneCompraDeuda(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_obtieneCompraDeuda '" & pcCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set obtieneCompraDeuda = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'CTI3 ERS0032020******************************************************************
Public Sub ReclasificaCategoriaCredito(ByVal pcCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec B2_RSE_stp_upd_Reclasificacion_cliente '" & pcCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
oConecta.Ejecutar (sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function DevuelveEstadosFinancieros(ByVal pcCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_RSE_sel_RecuperaEstadosFinanceros '" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveEstadosFinancieros = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub ActualizaCategoriaCredito(ByVal pcCtaCod As String, ByVal pcTpoCredito As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_RSE_upd_ActualizaTipoCredito '" & pcCtaCod & "','" & pcTpoCredito & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
oConecta.Ejecutar (sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function ObtieneTipoCredito(ByVal pcCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_RSE_sel_ObtieneTipoCredito '" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneTipoCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'********************************************************************************
'**ARLO20171127 FIN ERS070 - 2017
'**ARLO20180319 INICIO ERS070 - 2017 ANEXO 01
Public Function ListarCreditosPreDesemCompraDeuda(ByVal pcCtaCod As String, ByVal pnTipoAsig As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS070_2017_ListaPreDesemCompraDeuda '" & pcCtaCod & "'," & pnTipoAsig

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ListarCreditosPreDesemCompraDeuda = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Sub InsertarPreDesembCompraDeuda(ByVal pcCtaCod As String, ByVal nMontoApro As Double, nDesemAComprar As Double, dFecha As Date, _
                                        ByVal pnEstado As Integer, ByVal pcMovNro As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertarPreDesembCompraDeuda
   
   sSql = "Exec stp_ins_ERS070_2017_PreDesemCompraDeuda '" & pcCtaCod & "'," & nMontoApro & "," & nDesemAComprar & ",'" & Format(dFecha, "mm/dd/yyyy") & "'," & pnEstado & ", '" & pcMovNro & "'"
          
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertarPreDesembCompraDeuda:
    Err.Raise Err.Number, "Insertar Datos Comite", Err.Description
End Sub
Public Function VerificaPreDesemCompraDeuda(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS070_2017_VerificaPreDesemCompraDeuda '" & pcCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set VerificaPreDesemCompraDeuda = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'**ARLO20180319 FIN ERS070 - 2017 ANEXO 01


' *** CTI2 20181012 **
Public Function VerificaItemTramaBcoNac(ByVal psFechaProceso As String, ByVal psCtaCod As String, _
                                        ByVal pnCuota As Integer, ByVal nImporteCuota As Currency) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo VerificaItemTramaBcoNac
    
    sSql = " EXEC stp_sel_BNCobrosVerificaItem   '" & psFechaProceso & "', '" & _
            psCtaCod & "', " & pnCuota & ", " & nImporteCuota
    
    VerificaItemTramaBcoNac = False
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs.RecordCount > 0 Then
            VerificaItemTramaBcoNac = True
        End If
    End If
    
    Exit Function
VerificaItemTramaBcoNac:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function VerificaItemTramaBcoNacCorresponsalia(ByVal psFechaProceso As String, ByVal psCtaCod As String, _
                                        ByVal pnCuota As Integer, ByVal nImporteCuota As Currency) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo VerificaItemTramaBcoNacCorresponsalia
    
    sSql = " EXEC stp_sel_BNCorresponsaliaVerificaItem '" & psFechaProceso & "', '" & _
            psCtaCod & "', " & pnCuota & ", " & nImporteCuota
    
    VerificaItemTramaBcoNacCorresponsalia = False
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs.RecordCount > 0 Then
            VerificaItemTramaBcoNacCorresponsalia = True
        End If
    End If
    
    Exit Function
VerificaItemTramaBcoNacCorresponsalia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
' *** END CTI2 *******

'CTI2 20181205 Corrección de pagos ****************
Public Function VerificaItemTramaBCP(ByVal psFecha As String, ByVal psCtaCod As String, _
                                     ByVal pTotal As Double, _
                                     ByVal psOperacion As String, _
                                     ByVal psOficina As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset

On Error GoTo VerificaItemTramaBCP
    
    sSql = " EXEC stp_sel_BCP_VerificaItem '" & psFecha & "', '" & _
            psCtaCod & "', " & pTotal & ", '" & psOperacion & "', '" & psOficina & "'"
    
    VerificaItemTramaBCP = False
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs.RecordCount > 0 Then
            VerificaItemTramaBCP = True
        End If
    End If
    
    Exit Function
VerificaItemTramaBCP:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END CTI2 *********************

'JOEP20190322 +++++ Mejora en el Extorno de pago +++++
Public Sub UpdExtornoPagoNor(ByVal pcCtaCod As String, ByVal pcMovNro As String, nCuota As String, ByVal pdFechaSis As String, ByVal pcCodUser As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInseUpdSelExtornoPagoNor

   sSql = "Exec stp_upd_ActPagoExtorno '" & pcCtaCod & "','" & pcMovNro & "','" & nCuota & "','" & Format(pdFechaSis, "yyyymmdd hh:mm:ss") & "','" & pcCodUser & "'," & pnEstado & ""

       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql

       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInseUpdSelExtornoPagoNor:
    Err.Raise Err.Number, "No se inserto el registro del Pago ", Err.Description
End Sub
Public Function SelExtornoPagoNor(ByVal pcCtaCod As String, ByVal pcMovNro As String, ByVal nTipoOpc As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
On Error GoTo SelExtornoPagoNor
    
sSql = "Exec stp_sel_RegistraPagoExtorno '" & pcCtaCod & "','" & pcMovNro & "'," & nTipoOpc & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set SelExtornoPagoNor = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
    
Exit Function
SelExtornoPagoNor:
    Err.Raise Err.Number, "Error En SelExtornoPagoNor", Err.Description
End Function

Public Function ValidaExtornoPagoNor(ByVal pcCtaCod As String, ByVal pcMovNro As String, ByVal pnCuota As String, ByVal pdFechaSis As String, ByVal pcCodUser As String, ByVal pnMonto As Currency, ByVal pcCodAge As String, ByVal pnEstado As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
On Error GoTo ValidaExtornoPagoNor
    
sSql = "Exec stp_sel_ValidadExtornoPag '" & pcCtaCod & "','" & pcMovNro & "','" & pnCuota & "','" & Format(pdFechaSis, "yyyymmdd hh:mm:ss") & "','" & pcCodUser & "'," & pnMonto & ",'" & pcCodAge & "'," & pnEstado & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidaExtornoPagoNor = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
    
Exit Function
ValidaExtornoPagoNor:
    Err.Raise Err.Number, "Error En ValidaExtornoPagoNor", Err.Description
End Function
'JOEP20190322 +++++ Mejora en el Extorno de pago +++++


'JOEP20210511 Segmentacion Prendario Externo
Public Function DataSegPrendarioExterno(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
On Error GoTo DataSegPrendarioExterno
    
sSql = "Exec stp_sel_ObtenerDatoSegPrendarioExterno '" & pcPersCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set DataSegPrendarioExterno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
    
Exit Function
DataSegPrendarioExterno:
    Err.Raise Err.Number, "Error En SegPrendarioExterno", Err.Description
End Function

Public Function DataSegExperianExterno(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
On Error GoTo DataSegExperianExterno
    
sSql = "Exec stp_sel_ObtenerDatoSegExperianExterno '" & pcPersCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set DataSegExperianExterno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
    
Exit Function
DataSegExperianExterno:
    Err.Raise Err.Number, "Error En DataSegExperianExterno", Err.Description
End Function
'JOEP20210511 Segmentacion Prendario Externo

'JOEP20211013 ERS048-2021 Restrincion de Desembolso >3500
Public Function ResPago(ByVal pcCtaCod As String, ByVal pnMontoDes As Double, ByVal pnMontoApro As Double, ByVal pnTpOperacion As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
On Error GoTo ResPago
    
sSql = "Exec stp_sel_RestrincionPagoDesembolso '" & pcCtaCod & "'," & pnMontoDes & "," & pnMontoApro & ",'" & pnTpOperacion & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ResPago = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
    
Exit Function
ResPago:
    Err.Raise Err.Number, "Error En Restrincion de Pago", Err.Description
End Function
'JOEP20211013 ERS048-2021 Restrincion de Desembolso >3500
