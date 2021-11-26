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
    Optional ByVal pbRefin As Boolean = False) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim vsEstadoCred As String
Dim vsProductos As String
Dim I As Integer
Dim bSolicitado As Boolean

    On Error GoTo ErrorRecuperaPersonasEstadoCred
    vsEstadoCred = ""
    bSolicitado = False
    For I = 0 To UBound(pnEstadoCred)
        vsEstadoCred = vsEstadoCred & pnEstadoCred(I) & ","
        If CInt(pnEstadoCred(I)) = gColocEstSolic Then
            bSolicitado = True
        End If
    Next I
    vsEstadoCred = Mid(vsEstadoCred, 1, Len(vsEstadoCred) - 1)
    vsProductos = ""
    If IsArray(pMatProd) Then
        vsProductos = "('"
        For I = 0 To UBound(pMatProd)
            vsProductos = vsProductos & pMatProd(I) & "','"
        Next I
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
    If bSolicitado Then
        If pbRefin Then
            sSql = sSql & " AND  C.cCtaCod in (Select cCtaCod From ColocacRefinanc )"
        Else
            sSql = sSql & " AND  C.cCtaCod not in (Select cCtaCod From ColocacRefinanc )"
        End If
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
    sSql = sSql & " nCuotaPend = (Select Min(nCuota) as nCuota From ColocCalendario Where cCtaCod = P.cCtaCod And nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & "),"
    sSql = sSql & " dFecVencPend = (select dVenc from ColocCalendario where cCtaCod = P.cCtaCod AND nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & "),"
    sSql = sSql & " nCuota = (Select Min(nCuota) From ColocCalendario Where cCtaCod = P.cCtaCod And nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND nColocCalendApl = " & gColocCalendAplCuota & ")"
    sSql = sSql & " from Producto P Inner Join ProductoPersona PP "
    sSql = sSql & " ON P.cCtaCod = PP.cCtaCod "
    sSql = sSql & " Where PP.cPersCod = '" & psPersCod & "' AND nPrdPersRelac = " & gPersTipoConvenio
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
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
    sSql = " Select  P.cCtaCod,A.cAgeDescripcion, C.dApertura, CN.cConsDescripcion as cTipoAho, "
    sSql = sSql & " CN2.cConsDescripcion as cEstado, CN3.cConsDescripcion as cParticip,"
    sSql = sSql & " CN4.cConsDescripcion as cMotivo, "
    sSql = sSql & " sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' when '2' then 'DOLARES' end"
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " Inner Join Captaciones C ON C.cCtaCod = P.cCtaCod"
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
    sSql = sSql & " P.nSaldo , CT2.cConsDescripcion cParticip"
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

Public Function DatosPosicionCliente(ByVal psPersCod As String) As ADODB.Recordset
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
    sSql = sSql & " dCancelado = (select dPrdEstado from ColocacEstado Where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstCancelado & ")"
    sSql = sSql & " from Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod"
    sSql = sSql & " inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & " inner join Colocaciones CL ON P.cCtaCod = CL.cCtaCod"
    sSql = sSql & " left join Constante CT ON P.nPrdEstado = CT.nConsValor And CT.nConsCod =" & gColocEstado
    sSql = sSql & " left join Constante CT2 ON PP.nPrdPersRelac = CT2.nConsValor And CT2.nConsCod ='" & gColocRelacPers & "'"
    sSql = sSql & " left join RelCtaCred RC ON PP.cCtaCod = RC.cCodCta"
    sSql = sSql & " where PP.cPersCod = '" & psPersCod & "'"
    
    Conn.AbreConexion
    Set DatosPosicionCliente = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
    
ErrorDatosPosicionCliente:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Sub NuevoPermisoNivelApr(ByVal cCodNiv As String, ByVal cPerscod As String)
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorNuevoPermisoNivelApr
    
    sSql = "INSERT INTO ColocCredPersNivelesApr(cCodNiv,cPersCod)"
    sSql = sSql & " VALUES('" & cCodNiv & "','" & cPerscod & "')"
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
Dim I As Integer

    On Error GoTo ErrorRecuperaCreditosPagoLote
    sCadCodMod = ""
    For I = 0 To UBound(MatDatos) - 1
        sCadCodMod = sCadCodMod & MatDatos(I) & "'"
    Next I
    
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
    sSql = sSql & "             And CV.cCodModular in ('" & sCadCodMod & ") "
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
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
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



Public Function RecuperaCreditosPagoLote(ByVal psCodInst As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaCreditosPagoLote
    sSql = "Select ' ',P.cCtacod, CV.cCodModular, CC.nNroProxCuota nCuota, SUM(CD.nMonto) nPago, "
    sSql = sSql & "Pers.cPersNombre, SUM(CD.nMonto - CD.nMontoPagado) nSaldoCuota, "
    sSql = sSql & " nMora = (Select SUM(CD2.nMonto - CD2.nMontoPagado) From ColocCalendDet CD2 WHERE CD2.cCtaCod = P.cCtaCod AND CD2.nColocCalendApl = 1 AND CD2.nNroCalen = CC.nNroCalen  AND CD2.nPrdConCeptoCod = 1101), "
    sSql = sSql & "CC.nDiasAtraso, "
    sSql = sSql & " nDeudaTotal = (Select SUM(nMonto - nMontoPagado) From ColocCalendDet Where cCtaCod = P.cCtaCod AND nColocCalendApl = " & gColocCalendAplCuota & " AND nNroCalen = CC.nNroCalen), "
    'sSql = sSql & " SUM(CD.nMonto - CD.nMontoPagado) nDeudaTotal, "
    sSql = sSql & " CC.cMetLiquidacion, "
    sSql = sSql & "P.nTransacc, P.nPrdEstado "
    sSql = sSql & "From Producto P Inner Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocCalendDet CD ON CD.cCtaCod = P.cCtaCod AND CD.nColocCalendApl = " & gColocCalendAplCuota & " AND CD.nNroCalen = CC.nNroCalen AND CD.nCuota = CC.nNroProxCuota AND CD.nPrdConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    sSql = sSql & "                Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    'sSQL = sSQL & "                Inner Join ColocCalendDet CD3 ON CD3.cCtaCod = P.cCtaCod AND CD3.nColocCalendApl = " & gColocCalendAplCuota & " AND CD3.nNroCalen = CC.nNroCalen  AND CD3.nColocConCeptoCod <> " & gColocConceptoCodInteresMoratorio
    sSql = sSql & " WHERE CV.cPersCod = '" & psCodInst & "' And P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSql = sSql & " GROUP BY P.cCtacod, CV.cCodModular, CC.nNroProxCuota, Pers.cPersNombre, CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, P.nPrdEstado, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto - CD.nMontoPagado) > 0 "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPagoLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosPagoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosExtornoGeneral(ByVal pdFecOper As Date, ByVal pTipoExt As TipoExtorno, Optional psCodUsu As String = "", Optional psCtaCod As String = "", Optional psCodCli As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosExtornoGeneral
    Select Case pTipoExt
        Case gTipoExtornoCuenta
            sSql = "Select MC.cCtaCod, M.cMovNro, M.cMovDesc, MC.nMovNro, MC.nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora "
            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And MC.cCtaCod = '" & psCtaCod & "'"
            sSql = sSql & "         And MC.nMovNro in (Select MAX(nMovNro) From MovCol Where SUBSTRING(cOpeCod,1,3)='100' AND cOpeCod <> '100999' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & " Order By M.nMovNro DESC "
        Case gTipoExtornoUsuario
            sSql = "Select MC.cCtaCod, M.cMovDesc, MC.nMovNro, MC.nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora"
            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And Right(M.cMovNro,4) = '" & psCodUsu & "'"
            sSql = sSql & "         And MC.nMovNro in (Select MAX(nMovNro) From MovCol Where SUBSTRING(cOpeCod,1,3)='100' AND cOpeCod <> '100999' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & " Order By M.nMovNro DESC "
            
        Case gTipoExtornoCliente
            sSql = "Select MC.cCtaCod, M.cMovDesc, MC.nMovNro, MC.nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora "
            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & "     Inner Join ProductoPersona PP ON PP.cCtaCod =  MC.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' And PP.cPersCod = '" & psCodCli & "'"
            sSql = sSql & "         And MC.nMovNro in (Select MAX(nMovNro) From MovCol Where SUBSTRING(cOpeCod,1,3)='100' AND cOpeCod <> '100999' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & " Order By M.nMovNro DESC "
            
        Case gTipoExtornoGeneral
            sSql = "Select MC.cCtaCod, M.cMovDesc, MC.nMovNro, MC.nMonto, RIGHT(M.cMovNro,4) cUsuario, MC.cOpeCod, substring(M.cMovNro,9,2) + ':' + substring(M.cMovNro,11,2) + ':' + substring(M.cMovNro,13,2) cHora "
            sSql = sSql & " From MovCol MC inner join Mov M ON MC.nMovNro = M.nMovNro "
            sSql = sSql & " Where LEFT(M.cMovNro,8) = '" & Format(pdFecOper, "yyyymmdd") & "' "
            sSql = sSql & "         And MC.nMovNro in (Select MAX(nMovNro) From MovCol Where SUBSTRING(cOpeCod,1,3)='100' AND cOpeCod <> '100999' Group by cCtaCod) AND M.nMovFlag <> 2"
            sSql = sSql & " Order By M.nMovNro DESC "
            
    End Select
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosExtornoGeneral = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosExtornoGeneral:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
'By Capi 08022008 para obtener los datos  de CreditoConsol para stp_sel_ReporteDetCalCarteraxIntDevSupProv
'By Capi 17112008
'Public Function ObtenerCalCarteraxIntDevSupProv(ByVal psOpcion As String, ByVal pnPeriodo As Integer, ByVal pnTipoCambio As Double)
Public Function ObtenerCalCarteraxIntDevSupProv(ByVal psOpcion As String, ByVal pnPeriodo As Integer, ByVal pnTipoCambio As Double, Optional ByVal pnMes As Integer, Optional ByVal pnAnno As Integer)
    Dim sSql As String
    Dim oCon As New DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    'by capi 17112008
    'sSQL = " Exec stp_sel_ReporteDetCalCarteraxIntDevSupProv '" & psOpcion & "'," & pnPeriodo & "," & pnTipoCambio & ""
    sSql = " Exec stp_sel_ReporteDetCalCarteraxIntDevSupProv '" & psOpcion & "'," & pnPeriodo & "," & pnTipoCambio & "," & pnMes & "," & pnAnno
    '
    Set ObtenerCalCarteraxIntDevSupProv = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
 End Function

'By Capi 08022008 para obtener los datos  de CreditoConsol para stp_sel_ReporteResCalCarteraxLineas
Public Function ObtenerCalCarteraxLineas(ByVal psOpcion As String, ByVal pnPeriodo As Integer, ByVal pnTipoCambio As Double, Optional ByVal pnMes As Integer, Optional ByVal pnAnno As Integer)
    Dim sSql As String
    Dim oCon As New DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    'ByVal 17112008
    'sSQL = " Exec stp_sel_ReporteResCalCarteraxLineas '" & psOpcion & "'," & pnPeriodo & "," & pnTipoCambio & ""
    sSql = " Exec stp_sel_ReporteResCalCarteraxLineas '" & psOpcion & "'," & pnPeriodo & "," & pnTipoCambio & "," & pnMes & "," & pnAnno
    '
    Set ObtenerCalCarteraxLineas = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
 End Function
'ALPA 20101228*************************************************************
Public Function nExisteTabla(ByVal lsNombreTabla As String, ByVal psServConsol As String) As Boolean
Dim rs As New ADODB.Recordset
Dim oCon As DConecta
Dim lsSql As String

nExisteTabla = False
Set oCon = New DConecta
oCon.AbreConexion
lsSql = "Select * From " & psServConsol & "SysObjects Where Name = '" & lsNombreTabla & "' and type ='U'"
Set rs = oCon.CargaRecordSet(lsSql)
oCon.CierraConexion
If Not (rs.BOF And rs.EOF) Then
    nExisteTabla = True
End If
rs.Close
Set rs = Nothing
End Function
Public Function RecuperaDatosReporteVerificacionIntDevSuspen(pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorRecuperaDatosReporteVerificacionIntDevSuspen
    sFecha = Format(pdFecha, "YYYYMM")
    If Len(Trim(sFecha)) <> 6 Then
        Set RecuperaDatosReporteVerificacionIntDevSuspen = Nothing
        Exit Function
    End If
    If nExisteTabla("RCDvc" & sFecha & "01", "DBConsolidada..") = False Then
        Set RecuperaDatosReporteVerificacionIntDevSuspen = Nothing
        Exit Function
    End If
    'sSql = "exec B2_stp_sel_reporteVerificacionInteresesDevenSuspe"
    sSql = " select isnull(CP.cAgeCodAct,'') cAgeCod,A.cAgeDescripcion,"
    sSql = sSql & "  case substring(CP.cCtaCod,9,1) when '1' then 'SOLES' else 'DOLARES' end cMoneda,"
    sSql = sSql & "    CP.cCalGen,"
    sSql = sSql & "        Case CP.cCalGen"
    sSql = sSql & "             when '0' then '0 - Normal'"
    sSql = sSql & "             when '1' then '1 - CPP'"
    sSql = sSql & "             when '2' then '2 - Deficiente'"
    sSql = sSql & "             when '3' then '3 - Dudoso'"
    sSql = sSql & "             when '4' then '4 - Perdida'"
    sSql = sSql & "         end cCalGenDescr,"
    sSql = sSql & "         x.cPersCod,"
    sSql = sSql & "         P.cPersNombre,"
    sSql = sSql & "         CP.cRefinan,"
    sSql = sSql & "         CP.cTpoCredCod,"
    sSql = sSql & "         Cons.cConsDescripcion,"
    sSql = sSql & "         x.cCtaCod,"
    sSql = sSql & "         case when left(cCtaCnt,2)='81' then nSaldo else 0 end nIntSusp,"
    sSql = sSql & "         case when left(cCtaCnt,2)='14' then nSaldo else 0 end nIntDev"
    sSql = sSql & "  From"
    sSql = sSql & "     (   select *"
    sSql = sSql & "         from DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02"
    sSql = sSql & "         where cCtaCnt like '81[12]40[123]%'"
    sSql = sSql & "         Union"
    sSql = sSql & "         select *"
    sSql = sSql & "         from DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02"
    sSql = sSql & "         where cCtaCnt like '14[12]8%' "
    sSql = sSql & "      ) X "
    sSql = sSql & "         inner join Persona P on x.cPersCod=P.cPersCod"
    sSql = sSql & "         inner join ColocCalifProv CP on x.cCtaCod=CP.cCtaCod"
    sSql = sSql & "         inner join Agencias A on A.cAgeCod=CP.cAgeCodAct"
    sSql = sSql & "         inner join Constante Cons on CP.cTpoCredCod=Cons.nConsValor and Cons.nConsCod=3034"

Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaDatosReporteVerificacionIntDevSuspen = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosReporteVerificacionIntDevSuspen:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
'**************************************************************************
Public Function RecuperaDatosParaValidarInteresesDiferidos(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal pnTipoCambio As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorRecuperaDatosParaValidarInteresesDiferidos
    sSql = "exec B2_stp_sel_ValidacionDeAjusteInteresesDiferidos '" & Format(pdFecha, "YYYY/MM/DD") & "'," & psMoneda & "," & pnTipoCambio

Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaDatosParaValidarInteresesDiferidos = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosParaValidarInteresesDiferidos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaDatosParaMntCreditosCofide(ByVal psLineaCred As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorRecuperaDatosParaMntCreditosCofide
    sSql = "exec stp_sel_ListaDeCreditoCofidexCalificacion  '" & psLineaCred & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosParaMntCreditosCofide = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosParaMntCreditosCofide:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'ALPA 20110810
Public Function RecuperaDatosParaMntCreditosCofidexCredito(ByVal psMoneda As String, ByVal psPlazo As String, ByVal psProducto As String, ByVal pnTasa As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorRecuperaDatosParaMntCreditosCofidexCredito
    sSql = "exec stp_sel_ListaLineaCreditoxCredito '" & psMoneda & "','" & psPlazo & "','" & psProducto & "' , " & pnTasa

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosParaMntCreditosCofidexCredito = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosParaMntCreditosCofidexCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub ActualizaLineaCredito(ByVal pdFecha As Date, ByVal psCtaCod As String, ByVal psLineaCred As String, ByVal psLineaCredAnterior As String)
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorActualizaLineaCredito
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    'oConecta.ConexionActiva.BeginTrans
    sSql = "stp_ins_CreditoColocLineaCredito '" & Format(pdFecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & psLineaCred & "', '" & psLineaCredAnterior & "'"
    oConecta.ConexionActiva.Execute sSql
    sSql = "stp_upd_CambiarLineaCredito '" & psCtaCod & "' , '" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psLineaCred & "'"
    oConecta.ConexionActiva.Execute sSql
    'oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorActualizaLineaCredito:
    Err.Raise Err.Number, "Actualiza Linea de Credito COFIDE", Err.Description
End Sub

Public Function ReporteLineaCredito(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito
    sSql = "exec stp_sel_ListaCreditosxCambiosDeLinea '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaDatosReporteAdeudadosCalendario() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosReporteAdeudadosCalendario
    sSql = "exec stp_sel_ReporteAdeudadoCalendario"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosReporteAdeudadosCalendario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosReporteAdeudadosCalendario:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosReporteAdeudadosCalendarioVigente(ByVal dFecha As Date, ByVal psMoneda As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosReporteAdeudadosCalendarioVigente
    sSql = "exec stp_sel_ReporteAdeudadoCalendarioTesoriax15A '" & Format(dFecha, "YYYY/MM/DD") & "','" & psMoneda & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosReporteAdeudadosCalendarioVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosReporteAdeudadosCalendarioVigente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaDatosReporteAnalisisDeCtaCyLPlazo(ByVal dFecha As Date, ByVal sMoneda As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosReporteAnalisisDeCtaCyLPlazo
    sSql = "exec stp_sel_AnalisisCtaAdeCyLPlazo '" & Format(dFecha, "YYYY/MM/DD") & "','" & sMoneda & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosReporteAnalisisDeCtaCyLPlazo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosReporteAnalisisDeCtaCyLPlazo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'EJVG20111219 *********************************************************************
Public Function ObtenerVecesCreditoyAhorroPersona(ByVal psPersCod As String) As Integer
    Dim sSql As String
    Dim oConecta As DConecta
    Dim rs As ADODB.Recordset
    On Error GoTo ErrorObtenerVecesCreditoyAhorroPersona
    sSql = " exec stp_sel_CuentaCreditosAhorroPersona '" & psPersCod & "'"
    
    Set oConecta = New DConecta
    Set rs = New ADODB.Recordset
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    ObtenerVecesCreditoyAhorroPersona = rs!nVeces
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerVecesCreditoyAhorroPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteMIF(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteMIF
    sSql = "exec stp_sel_FormatoIFFoncodes '" & Format(pdFecha, "YYYY/MM/DD") & "', " & pnTipoCambio

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteMIF = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteMIF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MIOL 20121026, SEGUN RQ12338 ********************************************************************
Public Function ReporteLineaCredito_Personeria(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_Personeria
    sSql = "exec stp_sel_TipoPersoneriaxLineaCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_Personeria = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_Personeria:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteLineaCredito_PersoneriaMes(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_PersoneriaMes
    sSql = "exec stp_sel_TipoPersoneriaxColocacionesMesCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_PersoneriaMes = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_PersoneriaMes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function ReporteLineaCredito_SectorEconomico(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_SectorEconomico
    sSql = "exec stp_sel_SectorEconomicoxLineaCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_SectorEconomico = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_SectorEconomico:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteLineaCredito_SectorEconomicoMes(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_SectorEconomicoMes
    sSql = "exec stp_sel_SectorEconomicoxColocacionesMesCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_SectorEconomicoMes = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_SectorEconomicoMes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteLineaCredito_UbicaGeografica(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_UbicaGeografica
    sSql = "exec stp_sel_UbicaGeograficaxLineaCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_UbicaGeografica = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_UbicaGeografica:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteLineaCredito_UbicaGeograficaMes(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_UbicaGeograficaMes
    sSql = "exec stp_sel_UbicaGeograficaxColocacionesMesCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_UbicaGeograficaMes = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_UbicaGeograficaMes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteLineaCredito_Consolidado(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteLineaCredito_Consolidado
    sSql = "exec stp_sel_ReporteConsolCOFO '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteLineaCredito_Consolidado = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteLineaCredito_Consolidado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ReporteFinancieroFoncodes(ByVal pdFecha As Date, ByVal psLineaCredito As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sFecha As String
    
    On Error GoTo ErrorReporteFinancieroFoncodes
    sSql = "exec stp_sel_InfFinancieraCreditosVigentesxCalif '" & Format(pdFecha, "yyyyMMdd") & "','" & psLineaCredito & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteFinancieroFoncodes = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorReporteFinancieroFoncodes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END MIOL ****************************************************************************************
'ALPA20130708*************************************************************************************
Public Function ReporteValidacionComisionCF(ByVal dFecha As Date, ByVal sMoneda As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorReporteValidacionComisionCF
    sSql = "exec stp_sel_ReporteValidacionComisionCartaFianza '" & Format(dFecha, "YYYY/MM/DD") & "','" & sMoneda & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReporteValidacionComisionCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorReporteValidacionComisionCF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'*************************************************************************************************
