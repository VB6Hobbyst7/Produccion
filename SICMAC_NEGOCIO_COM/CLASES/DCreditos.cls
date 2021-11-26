VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCreditos"
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

Public Function RecuperaHistoriaCredPersona(ByVal psPersCod As String) As Variant
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As DConecta
Dim MatDatos() As String
Dim nPunt As Integer

    Set oConn = New DConecta
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
Dim oConecta As DConecta
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
    
    Set oConecta = New DConecta
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
    Optional ByVal pbRefin As Boolean = False, Optional pbMuestraTodos As Boolean = False) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
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
    sSql = sSql & " AND SubString(Prd.cCtaCod,4,2)='" & psCodAgencia & "'"
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPersonasEstadoCredAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaPersonasEstadoCred:
    
End Function

Public Sub ActualizaNivel(ByVal nNivel As Integer, ByVal cDescrip As String)
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorActualizaNivel
    sSql = "UPDATE ColocCredNivelesTipos SET cDescripcion = '" & cDescrip & "'"
    sSql = sSql & " WHERE nNivel = " & Trim(Str(nNivel))
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorNuevonivel
    sSql = "INSERT INTO ColocCredNivelesTipos(nNivel,cDescripcion)"
    sSql = sSql & " VALUES(" & Trim(Str(nNivel)) & ",'" & cDescrip & "')"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorEliminaNivel
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaNivAprPersona
    sSql = "Select cCodNiv, cPersCod from ColocCredPersNivelesApr Where cPersCod = '" & sPersCod & "'"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaNiveles
    sSql = "Select nNivel,cDescripcion from ColocCredNivelesTipos"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

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
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

    sSql = "Select * from ColocCredNivelesApr Where nNivel = " & Trim(psCodNiv)
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

    On Error GoTo ErrorExisteNivelAprobacion
    sSql = "Select * from ColocCredNivelesApr Where cCodNiv = '" & Trim(psCodNiv) & "' AND cProduct = '" & Trim(psProduct) & "' AND cTpoCred = '" & Trim(psTpoCred) & "' AND nNivel = " & pnNivel & " And nMoneda = " & pnMoneda
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaNivelesAprobacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaNivelesAprobacion:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion", Err.Description
End Function

Public Function DatosPosicionClienteAhorro(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As DConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteAhorro

    Set Conn = New DConecta
    sSql = " Select  P.cCtaCod,A.cAgeDescripcion, "
    sSql = sSql & " nSaldoCont= CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE P.nSaldo END, "
    sSql = sSql & " nSaldoDisp = CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE nSaldoDisp END, "
    sSql = sSql & " C.dApertura, CN.cConsDescripcion as cTipoAho, "
    sSql = sSql & " CN2.cConsDescripcion as cEstado, CN3.cConsDescripcion as cParticip,"
    sSql = sSql & " CN4.cConsDescripcion as cMotivo, "
    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' when '2' then 'DOLARES' end, cctacodant= case when R.CCTACODANT is null then '' else r.cctacodant end  "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " Inner Join Captaciones C ON C.cCtaCod = P.cCtaCod"
    sSql = sSql & " left join Relcuentas R ON R.cCTACOD=P.CCTACOD   "
    sSql = sSql & " Inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " Left join ProductoBloqueos PB ON PB.cCtaCod = P.cCtaCod AND PB.cMovNro = (Select Top 1 cMovNro From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC)"
    sSql = sSql & " LEFT Join Constante CN ON CONVERT(Int,SUBSTRING(P.cCtaCod,6,3))=CN.nConsValor AND CN.nConsCod = 1001"
    sSql = sSql & " LEFT join Constante CN2 ON CN2.nConsValor = P.nPrdEstado AND CN2.nConsCod = 2001"
    sSql = sSql & " LEFT join Constante CN3 ON CN3.nConsValor = PP.nPrdPersrelac AND CN3.nConsCod = 2005"
    sSql = sSql & " LEFT join Constante CN4 ON CN4.nConsValor = PB.nBlqMotivo AND CN4.nConsCod = 2007 "
    sSql = sSql & " WHERE PP.cpersCod = '" & psPersCod & "'"
    
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
Dim Conn As DConecta

    sSql = " Select P.cCtaCod, CR.dIngRecup, P.cCtaCod, A.cAgeDescripcion,"
    sSql = sSql & " cTipoCred=Case substring(P.cCtaCod,6,1) "
    sSql = sSql & "               when '1' then 'COMERCIAL'"
    sSql = sSql & "               when '2' then 'MICROEMPRESA'"
    sSql = sSql & "               when '3' then 'CONSUMO'"
    sSql = sSql & "               when '4' then 'HIPOTECARIO'"
    sSql = sSql & "               end,"
    sSql = sSql & " P.nPrdestado, CT.cConsDescripcion cEstado,"
    sSql = sSql & " CT2.cConsDescripcion as cParticip, P.nSaldo,"
    sSql = sSql & " dFecCast = (Select top 1 dPrdEstado from ColocacEstado where cctacod = P.cCtaCod AND nPrdEstado = " & gColocEstRecVigCast & "),"
    sSql = sSql & " cMovNro = (select top 1 cMovNro from Mov Where nMovnro = (select MAX(nMovnro) from MovCol Where cCtaCod = P.cCtaCod))"
    sSql = sSql & " From Producto P  Inner join ColocRecup CR ON P.cCtaCod = CR.cCtaCod"
    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod = " & gColocEstado
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod"
    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod = " & gColocRelacPers & " "
    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "'"
    
    Set Conn = New DConecta
    Conn.AbreConexion
    Set DatosPosicionClienteJudicial = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    

End Function


Public Function DatosPosicionClientePigno(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As DConecta

    sSql = "Select P.cCtacod, P.nPrdEstado, CT.cConsDescripcion cEstado, C.dVigencia,"
    sSql = sSql & " A.cAgeDescripcion, CP.nNroRenov, C.nMontoCol, C.dVenc, SUBSTRING(P.cCtaCod,4,2) as cAgeCod,"
    sSql = sSql & " P.nSaldo , CT2.cConsDescripcion cParticip,(case when substring(p.cctacod,9,1)='1' then 'SOLES' else 'DOLARES' end ) sMoneda ,"
    sSql = sSql & " CP.NTASACION "
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocPignoraticio CP  ON P.cCtaCod = CP.cCtaCod"
    sSql = sSql & " Left Join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
    sSql = sSql & " Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod"
    sSql = sSql & " Left Join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod"
    sSql = sSql & " Left Join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod =" & gColocRelacPers
    sSql = sSql & " Where PP.cPersCod = '" & psPersCod & "'"
    
    Set Conn = New DConecta
    Conn.AbreConexion
    Set DatosPosicionClientePigno = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    

End Function

Public Function DatosPosicionClientePignoLima(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim Conn As DConecta

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
    
    Set Conn = New DConecta
    Conn.AbreConexion
    Set DatosPosicionClientePignoLima = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

End Function

'Public Function DatosPosicionCliente(ByVal psPersCod As String) As ADODB.Recordset
'Dim Conn As DConecta
'Dim sSql As String
'
'On Error GoTo ErrorDatosPosicionCliente
'
'    Set Conn = New DConecta
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
'    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
'End Function

Public Function DatosPosicionCliente(ByVal psPersCod As String, Optional ByVal nMuestraAnalistas As Boolean = True) As ADODB.Recordset
Dim Conn As DConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionCliente

    Set Conn = New DConecta
    sSql = "Select P.cCtaCod,"
    sSql = sSql & " dSolicitado=(SELECT CE.dPrdEstado From ColocacEstado CE where  CE.cCtaCod = P.cCtaCod and CE.nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " A.cAgeDescripcion,"
    sSql = sSql & " cTipoCred=Case substring(P.cCtaCod,6,1)"
                              sSql = sSql & " when '1' then 'COMERCIAL'"
                              sSql = sSql & " when '2' then 'MICROEMPRESA'"
                              sSql = sSql & " when '3' then 'CONSUMO'"
                              sSql = sSql & " when '4' then 'HIPOTECARIO'"
                              sSql = sSql & " end,"
    sSql = sSql & " P.nPrdEstado, CT.cConsDescripcion as cEstadoDesc, PP.nPrdPersRelac, CT2.cConsDescripcion as cRelacionDesc,"
    sSql = sSql & " cPersAnalista = (select cUser From RRHH Where cPerscod = (Select top 1 cPersCod from ProductoPersona where nPrdPersRelac = '" & gColRelPersAnalista & "' and cCtaCod = P.cCtaCod)),"
    sSql = sSql & " nAnalistaNota = (Select nColocNota from ColocCalificacionAnalista where cCtaCod = P.cCtaCod and dColocNotaFecha = (select Max(dColocNotaFecha) from  ColocCalificacionAnalista where cCtaCod = P.cCtaCod )),"
    sSql = sSql & " CL.nMontoCol as nPrestamo, P.nSaldo, RC.cCodAnt1, RC.cCodAnt2,"
    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1)"
                    sSql = sSql & " when '1' then 'SOLES'"
                    sSql = sSql & " when '2' then 'DOLARES'"
                    sSql = sSql & " end,"
    'sSql = sSql & " dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & "),"
    sSql = sSql & " dCancelado = (select dPrdEstado from Producto Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & "),"
    sSql = sSql & " dAprobacion=(Select  CC1.dPago"
    sSql = sSql & " From ColocCalendario CC1"
    sSql = sSql & " Where CC1.cCtaCod=P.cCtaCod and CC1.nColocCalendApl=0 and CC1.nCuota=1"
    sSql = sSql & " and CC1.nNroCalen=(Select Max(nNroCalen) From ColocCalendario"
    sSql = sSql & " Where cCtaCod=CC1.cCtaCod and CC1.nColocCalendApl=nColocCalendApl and"
    sSql = sSql & " nCuota=CC1.nCuota and CC1.nColocCalendEstado=nColocCalendEstado)"
    sSql = sSql & " and CC1.nColocCalendEstado=1),"
    sSql = sSql & " cRFA=Case When CC.cRFA Is Null or CC.cRFA='NOR' Then 'NOR' Else CC.cRFA End"
    sSql = sSql & " from Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & " inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod"
    sSql = sSql & " left join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
    sSql = sSql & " left join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='" & gColocRelacPers & "'"
    sSql = sSql & " left join RelCtaCred RC ON PP.cCtaCod = RC.cCodCta"
    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "' AND SUBSTRING(P.cCtaCod,6,3) <> 121"
    
    If nMuestraAnalistas = False Then
        sSql = sSql & " and nprdpersrelac <>" & gColRelPersAnalista & "  and nprdpersrelac <>" & gColRelPersApoderado
    End If
    
    sSql = sSql & " Order By dSolicitado"
    
    Conn.AbreConexion
    Set DatosPosicionCliente = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    
ErrorDatosPosicionCliente:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Function DatosPosicionClienteCF(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As DConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteCF

    Set Conn = New DConecta
    sSql = "Select P.cCtaCod, dEmitido =(SELECT CF.dPrdEstado From ColocCFEstado CF " _
        & "WHERE CF.cCtaCod = P.cCtaCod and CF.nPrdEstado = 2020), A.cAgeDescripcion, P.nPrdEstado, " _
        & "CT.cConsDescripcion as cEstadoDesc, PP.nPrdPersRelac, CT2.cConsDescripcion as cRelacionDesc, " _
        & "cPersAnalista = (select cUser From RRHH  " _
        & "        Where cPerscod = (Select top 1 cPersCod from ProductoPersona WHERE nPrdPersRelac = '20' and cCtaCod = P.cCtaCod)), " _
        & "CL.nMontoCol as nPrestamo, P.nSaldo, sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' " _
        & "WHEN '2' then 'DOLARES' END, dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And " _
        & "         nPrdEstado = 2090) " _
        & " FROM Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod " _
        & " INNER JOIN Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod " _
        & "LEFT JOIN Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =3001 " _
        & "LEFT JOIN Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='3002' " _
        & "WHERE PP.cPersCod = '" & psPersCod & "' AND SUBSTRING(P.cCtaCod,6,3) IN (121, 221)"
        
    Conn.AbreConexion
    Set DatosPosicionClienteCF = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorDatosPosicionClienteCF:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Function DatosPosicionClienteComentarios(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As DConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteComentario

    Set Conn = New DConecta
    sSql = "Select C.cPersCod, C.cMovNro, C.cComentario From Persona P JOIN PersComentario C ON " _
        & "P.cPersCod = C.cPersCod WHERE P.cPersCod = '" & psPersCod & "' ORDER BY C.cMovNro DESC"
        
    Conn.AbreConexion
    Set DatosPosicionClienteComentarios = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorDatosPosicionClienteComentario:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Sub NuevoPermisoNivelApr(ByVal cCodNiv As String, ByVal cPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorNuevoPermisoNivelApr
    
    sSql = "INSERT INTO ColocCredPersNivelesApr(cCodNiv,cPersCod)"
    sSql = sSql & " VALUES('" & cCodNiv & "','" & cPersCod & "')"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarPermisoNivelApr
    sSql = "DELETE ColocCredPersNivelesApr WHERE cCodNiv = '" & cCodNiv & "' AND cPersCod = '" & psPersCod & "'"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorNuevoNivelAprobacion
    
    sSql = "INSERT INTO ColocCredNivelesApr(cCodNiv, cProduct, cTpoCred, nNivel, nMontoMax, nMontoMin, nMoneda)"
    sSql = sSql & " VALUES('" & cCodNiv & "','" & cProduct & "','" & cTpoCred & "'," & Trim(cNivel) & "," & Format(nMontoMax, "#0.00") & "," & Format(nMontoMin, "#0.00") & "," & pnMoneda & ")"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorEliminaNivelAprobacion
    
    Set oConecta = New DConecta
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

Public Function RecuperaCreditosArchivoPagoLote(ByVal MatDatos As Variant, ByVal psCodInst As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
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
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSql = sSql & "             And CV.cCodModular in (" & sCadCodMod & ") "
    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaCreditosCodigoModular
    sSql = "Select ' ',P.cCtacod, CV.cCodModular, "
    sSql = sSql & "Pers.cPersNombre "
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    'sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' "
    sSql = sSql & " AND SUBSTRING(P.cCtaCod,4,2) in " & psAgencia
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCodigoModular = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosCodigoModular:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Public Function RecuperaCreditosPagoLote(ByVal psCodInst As String, Optional ByVal pnMoneda As Moneda = gMonedaNacional) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaCreditosPagoLote
  sSql = " SELECT * FROM ( "
    sSql = sSql & " Select ' ' as cOK,P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, "
    sSql = sSql & " ROUND(SUM(CD.nMonto-CD.nMontoPagado),6) nPago,"
    sSql = sSql & "Pers.cPersNombre, ROUND(SUM(CD.nMonto - CD.nMontoPagado),2) nSaldoCuota, "
    sSql = sSql & " nMora = (Select ROUND(SUM(CD2.nMonto - CD2.nMontoPagado),2) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
    sSql = sSql & " CC.nDiasAtraso, "
    sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen ), "
    sSql = sSql & " CC.cMetLiquidacion, "
    sSql = sSql & "P.nTransacc, P.nPrdEstado, CC.nNroCalen, "
    sSql = sSql & "ROUND(dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) as nITF,   "
    sSql = sSql & " nCapital = (Select ROUND(SUM(CD3.nMonto - CD3.nMontoPagado),2) From ColocCalendDet CD3 WHERE CD3.cCtaCod = P.cCtaCod AND CD3.nColocCalendApl = 1 AND CD3.nNroCalen = CC.nNroCalen  AND CD3.nPrdConCeptoCod = 1000 AND CD3.nCuota = CC.nNroProxCuota), "
    sSql = sSql & " nInteres = (Select ROUND(SUM(CD4.nMonto - CD4.nMontoPagado),2) From ColocCalendDet CD4 WHERE CD4.cCtaCod = P.cCtaCod AND CD4.nColocCalendApl = 1 AND CD4.nNroCalen = CC.nNroCalen  AND CD4.nPrdConCeptoCod = 1100 AND CD4.nCuota = CC.nNroProxCuota), "
    sSql = sSql & " Protesto = (Select ISNULL(ROUND(SUM(CD6.nMonto - CD6.nMontoPagado),2),0) From ColocCalendDet CD6 WHERE CD6.cCtaCod = P.cCtaCod AND CD6.nColocCalendApl = 1 AND CD6.nNroCalen = CC.nNroCalen  AND CD6.nPrdConCeptoCod IN (1219,1220) and CD6.nCuota = CC.nNroProxCuota ), "
    sSql = sSql & " Gastos   = (Select ISNULL(ROUND(SUM(CD5.nMonto - CD5.nMontoPagado),2),0) From ColocCalendDet CD5 WHERE CD5.cCtaCod = P.cCtaCod AND CD5.nColocCalendApl = 1 AND CD5.nNroCalen = CC.nNroCalen  AND CD5.nPrdConCeptoCod not in (1000,1100,1101,1219,1220) AND CD5.nCuota = CC.nNroProxCuota ) "
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "'  And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") and SUBSTRING(P.CCTACOD,9,1)='" & pnMoneda & "' "
    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto - CD.nMontoPagado) > 0  ) AS PAGOSLOTE "
    sSql = sSql & " ORDER BY cPersNombre "
        
'    sSql = "Select ' ',P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, "
'    sSql = sSql & " ROUND(SUM(CD.nMonto-CD.nMontoPagado) + dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) nPago,"
'    sSql = sSql & "Pers.cPersNombre, ROUND(SUM(CD.nMonto - CD.nMontoPagado),2) nSaldoCuota, "
'    sSql = sSql & " nMora = (Select ROUND(SUM(CD2.nMonto - CD2.nMontoPagado),2) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
'    sSql = sSql & " CC.nDiasAtraso, "
'    sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen ), "
'    sSql = sSql & " CC.cMetLiquidacion, "
'    sSql = sSql & "P.nTransacc, P.nPrdEstado, CC.nNroCalen, "
'    sSql = sSql & "ROUND(dbo.ITFImpuesto(SUM(CD.nMonto - CD.nMontoPagado)),6) as nITF,   "
'    sSql = sSql & " nCapital = (Select ROUND(SUM(CD3.nMonto - CD3.nMontoPagado),2) From ColocCalendDet CD3 WHERE CD3.cCtaCod = P.cCtaCod AND CD3.nColocCalendApl = 1 AND CD3.nNroCalen = CC.nNroCalen  AND CD3.nPrdConCeptoCod = 1000 AND CD3.nCuota = CC.nNroProxCuota), "
'    sSql = sSql & " nInteres = (Select ROUND(SUM(CD4.nMonto - CD4.nMontoPagado),2) From ColocCalendDet CD4 WHERE CD4.cCtaCod = P.cCtaCod AND CD4.nColocCalendApl = 1 AND CD4.nNroCalen = CC.nNroCalen  AND CD4.nPrdConCeptoCod = 1100 AND CD4.nCuota = CC.nNroProxCuota), "
'    sSql = sSql & " Protesto = (Select ISNULL(ROUND(SUM(CD6.nMonto - CD6.nMontoPagado),2),0) From ColocCalendDet CD6 WHERE CD6.cCtaCod = P.cCtaCod AND CD6.nColocCalendApl = 1 AND CD6.nNroCalen = CC.nNroCalen  AND CD6.nPrdConCeptoCod IN (1219,1220) and CD6.nCuota = CC.nNroProxCuota ), "
'    sSql = sSql & " Gastos   = (Select ISNULL(ROUND(SUM(CD5.nMonto - CD5.nMontoPagado),2),0) From ColocCalendDet CD5 WHERE CD5.cCtaCod = P.cCtaCod AND CD5.nColocCalendApl = 1 AND CD5.nNroCalen = CC.nNroCalen  AND CD5.nPrdConCeptoCod not in (1000,1100,1101,1219,1220) AND CD5.nCuota = CC.nNroProxCuota ) "
'    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
'    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
'    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
'    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
'    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
'    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "'  And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") and SUBSTRING(P.CCTACOD,9,1)='" & pnMoneda & "' "
'    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
'    sSql = sSql & " HAVING SUM(CD.nMonto - CD.nMontoPagado) > 0  "
'    sSql = sSql & " ORDER BY  Pers.cPersNombre"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPagoLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosPagoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosExtornoGeneral(ByVal pdFecOper As Date, ByVal pTipoExt As TipoExtorno, Optional psCodUsu As String = "", Optional psctacod As String = "", Optional psCodCli As String = "", _
    Optional ByVal pbDesemb As Boolean = False, Optional ByVal psCodAge As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosExtornoGeneral
    Select Case pTipoExt
        Case gTipoExtornoCuenta
                   sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And MC.cCtaCod = '" & psctacod & "'"
            sSql = sSql & "       And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MovDiario M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
            'sSql = sSql & "       AND MC.cOpeCod <> '100301' "
            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
            If pbDesemb Then
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
            Else
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
            End If
            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
            sSql = sSql & " Order By M.nMovNro DESC "

        Case gTipoExtornoUsuario
            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago"
            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And Right(M.cMovNro,4) = '" & psCodUsu & "'"
            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOVDIARIO M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & "         AND MC.cOpeCod <> '100301' "
            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
            If pbDesemb Then
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
            Else
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
            End If
            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
            sSql = sSql & " Order By M.nMovNro DESC "
            
        Case gTipoExtornoCliente
            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
            sSql = sSql & " From MovCol MC inner join MovDiario M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & "     Inner Join ProductoPersona PP ON PP.cCtaCod =  MC.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And PP.cPersCod = '" & psCodCli & "'"
            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOVDIARIO M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & "         AND MC.cOpeCod <> '100301' "
            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
            If pbDesemb Then
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
            Else
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
            End If
            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
            sSql = sSql & " Order By M.nMovNro DESC "
            
        Case gTipoExtornoGeneral
            sSql = "Select O.cOpeDesc, SUM(MC.nSaldoCap) as nSaldoCap, MC.cCtaCod, M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) as nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora,MC.nPrepago "
            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Left Join OpeTpo O ON O.cOpeCod = MC.cOpeCod "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' "
            sSql = sSql & "         And MC.nMovNro in (Select MAX(MC1.nMovNro) From MovCol MC1 JOIN MOV M1 ON M1.NMOVNRO =MC1.NMOVNRO Where SUBSTRING(mc1.cOpeCod,1,3)='100' AND mc1.cOpeCod <> '100999' AND M1.NMOVFLAG<>'2' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & "         AND MC.cOpeCod <> '100301' "
            sSql = sSql & "         AND SUBSTRING(MC.cCtaCod,4,2)='" & psCodAge & "' "
            If pbDesemb Then
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) = '1001' "
            Else
                sSql = sSql & "       AND SUBSTRING(MC.cOpeCod,1,4) <> '1001' "
            End If
            sSql = sSql & " Group by O.cOpeDesc, MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro,MC.cOpeCod,M.nMovNro,MC.nPrepago "
            sSql = sSql & " Order By M.nMovNro DESC "
            
    End Select
    
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    oConecta.ConexionActiva.CommandTimeout = 60000
    
    Set RecuperaDatosExtornoGeneral = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosExtornoGeneral:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaPagosEnLote(ByVal pdFecOper As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "Select M.cMovDesc, MC.nMovNro, SUM(MC.nMonto) nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora "
    sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
    sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' "
    sSql = sSql & "         And M.cOpeCod =  '" & gCredPagLote & "' And MC.nMovNro in (Select MAX(MC2.nMovNro) From MovCol MC2 Inner join Mov M2 On  MC2.nMovNro = M2.nMovNro Where MC2.cCtaCod = MC.cCtaCod and MC2.copeCod not in ('" & gCredExtPago & "','" & gCredExtPagoLote & "')  And M2.nMovFlag = 0  Group by MC2.cCtaCod) "
    sSql = sSql & "         And M.nMovFlag = " & gMovFlagVigente
    sSql = sSql & " group by M.cMovDesc,MC.nMovNro,RIGHT(M.cMovNro,4),MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2)"
    sSql = sSql & " Order By MC.nMovNro DESC "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagosEnLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosdePagoEnLote(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro & " and NOT cOpecod LIKE '99%'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosdePagoEnLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function SaldoConsolCapta(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As Recordset
Dim oConecta As DConecta

    sSql = "SELECT SUM(nSaldo) as SaldoCaptac FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
        & "WHERE SUBSTRING(P.cCtaCod,6,3) IN (232,233,234,239) " _
        & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")" _
        & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 10 "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolCapta = IIf(IsNull(rs!SaldoCaptac), 0, rs!SaldoCaptac)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function SaldoConsolColoc(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As Recordset
Dim oConecta As DConecta

    sSql = "SELECT SUM(nSaldo) SaldoColoc FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
        & "WHERE SUBSTRING(P.cCtaCod,6,3) NOT IN (232,233,234,239,121,221) " _
        & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado NOT IN (" & gColocEstSolic & "," & gColocEstSug _
        & "," & gColocEstAprob & "," & gColocEstCancelado & "," & gColocEstRech & "," & gColocEstRefinanc & "," & gColocEstRetirado _
        & "," & gColPEstRegis & "," & gColPEstDifer & "," & gColPEstCance & "," & gColPEstRemat & "," & gColPEstAdjud _
        & "," & gColPEstSubas & "," & gColPEstAnula & "," & gColPEstChafa & "," & gColPEstAnNoD & "," & gPigEstRegis _
        & "," & gPigEstCancelPendRes & "," & gPigEstRescate & "," & gPigEstRematPFact & "," & gPigEstRematFact & " " _
        & "," & gPigEstAdjud & "," & gPigEstAnula & ")" _
        & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 20 "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    SaldoConsolColoc = IIf(IsNull(rs!SaldoColoc), 0, rs!SaldoColoc)
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function SaldoConsolCF(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Currency
Dim sSql As String
Dim rs As Recordset
Dim oConecta As DConecta

    sSql = "SELECT SUM(nSaldo) SaldoCF FROM PRODUCTO P INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD = PP.CCTACOD " _
        & "WHERE SUBSTRING(P.cCtaCod,6,3) IN (121,221)" _
        & "AND substring(P.cCtaCod, 9,1) = " & pnMoneda & " AND nPrdEstado IN (" & gColocEstVigNorm & ")" _
        & "AND cPersCod = '" & psPersCod & "' and PP.nPrdPersRelac = 20"

    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    On Error GoTo ErrorActualizaNivel
    sSql = "INSERT PersComentario (cPersCod,cMovNro,cComentario) " _
        & "Values ('" & sPersCod & "','" & sMovNro & "','" & sComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizaNivel:
    Err.Raise Err.Number, "Agrega Persona Comentario", Err.Description
End Sub








