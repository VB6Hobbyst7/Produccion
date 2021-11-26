VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaFiltroAplicadoCuenta(ByVal psTipoFiltro As String, ByVal bOtroFiltro As Boolean, _
Optional ByVal psOtraInstitucion = "") As ADODB.Recordset
Dim sSQL As String
Dim oConn As DConecta

    sSQL = "Select CF.nPrdConceptoCod, CF.nProdCod, "
    sSQL = sSQL & " CF.cAgeCod,CF.cTpoFiltro,CF.cIntitucion "
    sSQL = sSQL & " From ProductoConceptoFiltro CF "
    sSQL = sSQL & " Where CF.cTpoFiltro = '" & psTipoFiltro & "' "
    If bOtroFiltro = True Then
        sSQL = sSQL & " AND cIntitucion is NULL"
    End If
   ' sSQL = sSQL & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    
    If Len(psOtraInstitucion) > 0 Then
        'adiciono el gasto solo de esa institucion
    sSQL = sSQL & " UNION"
    sSQL = sSQL & " Select CF.nPrdConceptoCod, CF.nProdCod, "
    sSQL = sSQL & " CF.cAgeCod,CF.cTpoFiltro,CF.cIntitucion "
    sSQL = sSQL & " From ProductoConceptoFiltro CF "
    sSQL = sSQL & " Where CF.cTpoFiltro = '" & psTipoFiltro & "' "
    sSQL = sSQL & " AND cIntitucion='" & psOtraInstitucion & "'"
    'sSQL = sSQL & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    End If
    sSQL = sSQL & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    Set oConn = New DConecta
    oConn.AbreConexion
    Set RecuperaFiltroAplicadoCuenta = oConn.CargaRecordSet(sSQL)
    oConn.CierraConexion
    Set oConn = Nothing

End Function


Public Function ObtenerCodInstitucionByCredito(ByVal psCtaCod As String) As String
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select cPersCod From ColocacConvenio Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerCodInstitucionByCredito = rs!cPersCod
    End If
    Set rs = Nothing
End Function


Public Function RecuperaFiltros(ByVal pnConceptoCod As Long, ByVal psTpoFiltro As String) As ADODB.Recordset
Dim sSQL As String
Dim oConn As DConecta

    sSQL = " Select CF.nProdCod,  C.cConsDescripcion, "
    sSQL = sSQL & " CF.cAgeCod , A.cAgeDescripcion "
    sSQL = sSQL & " from ProductoConceptoFiltro CF"
    sSQL = sSQL & " INNER join ProductoConcepto PC ON CF.nPrdConceptoCod = PC.nPrdConceptoCod"
    sSQL = sSQL & " INNER Join Constante C ON C.nConsValor = CF.nProdCod AND C.nConsCod=" & IIf(psTpoFiltro = "G", "1027", "1001")
    sSQL = sSQL & " INNER Join Agencias A ON CF.cAgeCod = A.cAgeCod"
    sSQL = sSQL & " Where CF.nPrdConceptoCod = " & pnConceptoCod & " AND cTpoFiltro = '" & psTpoFiltro & "' "
    sSQL = sSQL & " Order by CF.nProdCod, CF.cAgeCod"
    
    Set oConn = New DConecta
    oConn.AbreConexion
    Set RecuperaFiltros = oConn.CargaRecordSet(sSQL)
    oConn.CierraConexion
    Set oConn = Nothing
End Function

Public Sub ActualizaGastosProdFiltro(ByVal pMatDatos As Variant, ByVal psTipoFiltro As String)
Dim oCred As DCredActualizaBD
Dim i As Integer
Dim sSQL As String
    Set oCred = New DCredActualizaBD
    oCred.dBeginTrans
    Call oCred.dDeleteProductoConceptoFitroTotal(CLng(pMatDatos(i, 0)), psTipoFiltro)
    For i = 0 To UBound(pMatDatos) - 1
        Call oCred.dInsertProductoConceptoFitro(CLng(pMatDatos(i, 0)), _
                        CInt(pMatDatos(i, 1)), pMatDatos(i, 5), psTipoFiltro, False, CStr(pMatDatos(i, 7)))
    Next i
    oCred.dCommitTrans
    Set oCred = Nothing
End Sub

Public Function RecuperaAgencias() As ADODB.Recordset
Dim oConn As DConecta
Dim sSQL As String

    Set oConn = New DConecta
    oConn.AbreConexion
    sSQL = "select cAgeCod, cAgeDescripcion From Agencias Order by cAgeCod "
    Set RecuperaAgencias = oConn.CargaRecordSet(sSQL)
    oConn.CierraConexion
    Set oConn = Nothing
End Function

Public Function RecuperaGastosCabecera(ByVal pnGastosProd As GastosProd) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta
Dim lsCondicionGastos As String
    If pnGastosProd = gGastosProdCredito Then
        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '12%'   "
    ElseIf pnGastosProd = gGastosProdJudicial Then
        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '32%'   "
    End If
    
    On Error GoTo ErrorRecuperaGastosCabecera
    
    sSQL = "Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor, "
    sSQL = sSQL & " nValor ,nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc,  "
    sSQL = sSQL & " cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc "
    sSQL = sSQL & " From ProductoConcepto "
    sSQL = sSQL & " Where nColocCred = " & pnGastosProd & " AND nPrdConceptoCod <> 1299 " & lsCondicionGastos
    sSQL = sSQL & " Order by convert(varchar(25),nPrdConceptoCod) "
   
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosCabecera = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosCabecera:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description

End Function

Public Function ExisteGasto(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
    ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, _
    ByVal pnConcepto As Long) As Boolean
    
    Dim sSQL As String
    Dim oCon As DConecta
    Dim R As ADODB.Recordset
    
    sSQL = "Select * From ColocCalendDet Where cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " AND nNroCalen = " & pnNroCalen
    sSQL = sSQL & " AND nColocCalendApl = " & pnColocCalendApl
    sSQL = sSQL & " AND nCuota = " & pnCuota
    sSQL = sSQL & " AND nPrdConceptoCod = " & pnConcepto
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    If R.RecordCount > 0 Then
        ExisteGasto = True
    Else
        ExisteGasto = False
    End If
    R.Close
    Set R = Nothing
End Function

Public Function RecuperaGastosAplicablesDesembolso(ByVal pnMoneda As Integer, _
    Optional ByVal psFiltro As String = "", Optional ByVal psTipoProceso As String = "'DE'", _
    Optional ByVal psGastoFijoVar As String = "N") As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaGastosAplicablesDesembolso
    
    If psTipoProceso = "DE" Then psTipoProceso = "'DE'"
    
    sSQL = "Select * from ProductoConcepto where nInicial >0 AND "
    sSQL = sSQL & " nAplicado in (" & gColocConceptoAplDesembolso & "," & gColocConceptoAplTodosD & "," & gColocConceptoAplTodosDC & ") "
    'sSql = sSql & " AND nPrdConceptoCod like '12%' AND (cFiltro = 'N' or cFiltro = '" & psFiltro & "') AND nMoneda = " & pnMoneda
    sSQL = sSQL & " AND nPrdConceptoCod like '12%' AND nMoneda = " & pnMoneda
    sSQL = sSQL & " AND ( nDeCierre = 0 OR nDeCierre is NULL) AND nColocCred = 1 AND cAplicaProceso in (" & psTipoProceso & ") "
    sSQL = sSQL & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
    
    If pnMoneda <> -1 Then
        sSQL = sSQL & " AND nMoneda = " & pnMoneda
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesDesembolso:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function

Public Function RecuperaGastosAplicablesCuotas(Optional ByVal pnMoneda As Integer = -1, _
    Optional ByVal psFiltroProceso As String = "", Optional ByVal psGastoFijoVar As String = "F", _
    Optional ByVal pbGastosCierre As Boolean = False, Optional ByVal MatGastosExon As Variant = "") As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
Dim i As Integer

    If Mid(psFiltroProceso, 1, 1) <> "'" Then
        psFiltroProceso = "'" & psFiltroProceso & "'"
    End If
    On Error GoTo ErrorRecuperaGastosAplicablesCuotas
    sSQL = "Select * from ProductoConcepto "
    sSQL = sSQL & " where nValor <> 0 AND nFinal >0 AND nAplicado in (" & gColocConceptoAplCuota & "," & gColocConceptoAplTodosC & "," & gColocConceptoAplTodosDC & ") "
    If pbGastosCierre Then
        sSQL = sSQL & " AND (nDeCierre = 1 or nDeCierre is null) "
    Else
        sSQL = sSQL & " AND (nDeCierre = 0 or nDeCierre is null) "
    End If
    sSQL = sSQL & " AND nColocCred = 1 AND nPrdConceptoCod like '12%' "
    If psFiltroProceso <> "" Then
        sSQL = sSQL & " AND cAplicaProceso in (" & IIf(psFiltroProceso = "'CA'" Or psFiltroProceso = "'PP'" Or psFiltroProceso = "'PA'", "'PA','CA','PP'", psFiltroProceso) & ")"
    End If
    sSQL = sSQL & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
    
    If pnMoneda <> -1 Then
        sSQL = sSQL & " AND nMoneda = " & pnMoneda
    End If
    
    If IsArray(MatGastosExon) Then
        If UBound(MatGastosExon) > 0 Then
            sSQL = sSQL & " AND nPrdConceptoCod Not In ("
           For i = 0 To UBound(MatGastosExon) - 1
                sSQL = sSQL & MatGastosExon(i) & ","
           Next i
           sSQL = Mid(sSQL, 1, Len(sSQL) - 1)
           sSQL = sSQL & ") "
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesCuotas = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesCuotas:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function


Public Function UltimoCorrelXProd() As Long
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

    sSQL = "Select MAX(convert(varchar(25),nPrdConceptoCod)) as nUlt From ProductoConcepto where nPrdConceptoCod like '12%' and nPrdConceptoCod <> 1299 "
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    UltimoCorrelXProd = IIf(IsNull(R!nUlt), 0, R!nUlt)
    UltimoCorrelXProd = UltimoCorrelXProd
    
End Function

Public Sub NuevoGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
    psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
    ByVal pnDiasApl As Integer, ByVal pnGastoTipoProd As Integer, ByVal pnEdadOper As Integer, _
    ByVal psGastoFijoVar As String, ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer)
    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorNuevoGasto
    sSQL = "INSERT INTO ProductoConcepto(nPrdConceptoCod, cDescripcion, nAplicado, nInicial, "
    sSQL = sSQL & " nFinal, nTpoValor, nValor, nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,"
    sSQL = sSQL & " cAplicaProceso, nOperador, nOperPorc,cOperMonto,nEdad,nDiasApl,nColocCred,nEdadOper,cGastoFijoVar,nOperDiasVenc,nDiasVenc, bSeguro)"
    sSQL = sSQL & " VALUES(" & psColocConceptoCod & ",'" & psDescripcion & "'," & psAplicado & "," & Format(pnInicial, "#0.00") & "," & Format(pnFinal, "#0.00") & "," & psTpoValor & "," & Format(pnValor, "#0.0000") & "," & psMoneda & "," & Format(pnMontoMin, "#0.00") & "," & Format(pnMontoMax, "#0.00") & ",'" & psAplMonto & "','" & psAplGarTpo & "',"
    sSQL = sSQL & "'" & psAplicaproceso & "'," & IIf(pnOperador = -1, "NULL", pnOperador) & "," & IIf(pnOperador = -1, "NULL", Format(pnOperPorc, "#0.00")) & "," & IIf(pnOperador = -1, "NULL", "'" & psOperMonto & "'") & "," & pnEdad & "," & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & "," & pnGastoTipoProd & "," & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ",'" & psGastoFijoVar & "'," & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & "," & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & "," & pnSeguro & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoGasto:
    Err.Raise Err.Number, "Nuevo Gasto", Err.Description
End Sub

Public Sub ActualizarGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
    ByVal psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
    ByVal pnDiasApl As Integer, ByVal pnEdadOper As Integer, ByVal psGastoFijoVar As String, _
    ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer)
Dim sSQL As String
Dim oConecta As DConecta
        
    
    
    On Error GoTo ErrorActualizarGasto
    sSQL = " UPDATE ProductoConcepto SET "
    sSQL = sSQL & "cDescripcion = '" & psDescripcion & "',"
    sSQL = sSQL & "nAplicado = " & psAplicado & ","
    sSQL = sSQL & "nInicial = " & Format(pnInicial, "#0.00") & ","
    sSQL = sSQL & "nFinal = " & Format(pnFinal, "#0.00") & ","
    sSQL = sSQL & "nTpoValor = " & psTpoValor & ","
    sSQL = sSQL & "nValor = " & Format(pnValor, "#0.0000") & ","
    sSQL = sSQL & "nMoneda = " & psMoneda & ", "
    sSQL = sSQL & " nMontoMin =  " & Format(pnMontoMin, "#0.00") & ","
    sSQL = sSQL & " nMontoMax =  " & Format(pnMontoMax, "#0.00") & ","
    sSQL = sSQL & " cAplicaMonto =  '" & psAplMonto & "',"
    sSQL = sSQL & " cFiltro = '" & psAplGarTpo & "', "
    sSQL = sSQL & " cAplicaProceso = '" & psAplicaproceso & "',"
    sSQL = sSQL & " nOperador = " & IIf(pnOperador = -1, "NULL", pnOperador) & ","
    sSQL = sSQL & " nOperPorc = " & IIf(pnOperPorc = -1, "NULL", Format(pnOperPorc, "#0.00")) & ","
    sSQL = sSQL & " cOperMonto = '" & psOperMonto & "',"
    sSQL = sSQL & " nEdad = " & pnEdad & ","
    sSQL = sSQL & " nEdadOper = " & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ","
    sSQL = sSQL & " nDiasApl = " & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & ","
    sSQL = sSQL & " cGastoFijoVar = '" & psGastoFijoVar & "', "
    sSQL = sSQL & " nOperDiasVenc = " & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & ", "
    sSQL = sSQL & " nDiasVenc = " & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & ","
    sSQL = sSQL & " bSeguro = " & pnSeguro
    sSQL = sSQL & " WHERE nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizarGasto:
    Err.Raise Err.Number, "Actualizar Gasto", Err.Description
End Sub

Public Function GastoAsignadoACredito(ByVal psColocConceptoCod As String) As Boolean
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    
    sSQL = "Select * From ColocCalendDet Where nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If R.RecordCount > 0 Then
        GastoAsignadoACredito = True
        R.Close
        Exit Function
        oConecta.CierraConexion
    Else
        GastoAsignadoACredito = False
    End If
    R.Close
    
    sSQL = "Select * From MovColDet Where nPrdConceptoCod = " & psColocConceptoCod
    Set R = oConecta.CargaRecordSet(sSQL)
    If R.RecordCount > 0 Then
        GastoAsignadoACredito = True
    Else
        GastoAsignadoACredito = False
    End If
    R.Close
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub EliminarGasto(ByVal psColocConceptoCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarGasto
    sSQL = " DELETE ProductoConcepto "
    sSQL = sSQL & " WHERE nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarGasto:
    Err.Raise Err.Number, "Eliminar Gasto", Err.Description
End Sub

Public Function RecuperaGastoValorPenalidadCancelacion(ByVal pnMoneda As Integer) As Double
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    On Error GoTo ErrorRecuperaGastoValorPenalidadCancelacion
    If pnMoneda = gMonedaNacional Then
        sSQL = "Select nValor from ProductoConcepto where nPrdConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoSoles
    Else
        sSQL = "Select nValor from ProductoConcepto where nPrdConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoDolares
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        RecuperaGastoValorPenalidadCancelacion = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
    Else
        RecuperaGastoValorPenalidadCancelacion = 0
    End If
    Exit Function

ErrorRecuperaGastoValorPenalidadCancelacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecupNomInstitucion(ByVal psCodInstitucion As String) As String
    Dim sSQL As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSQL = "Select cPersNombre From Persona Where cPersCod='" & psCodInstitucion & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        RecupNomInstitucion = rs!cpersnombre
    End If
    Set rs = Nothing
End Function


Public Function RecupCodInstitucion(ByVal psCodProductoConcepto As String, ByVal sProducto As String, ByVal psAgenciaCod As String) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    
    sSQL = "Select Pers.cPersCod,Pers.cPersNombre"
    sSQL = sSQL & " From PRODUCTOCONCEPTOFILTRO PCF"
    sSQL = sSQL & " Inner Join  Persona Pers on Pers.cPersCod=PCF.cIntitucion"
    sSQL = sSQL & " Where nPrdConceptoCod=" & psCodProductoConcepto & " and nProdCod=" & sProducto & " and cAgeCod='" & psAgenciaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecupCodInstitucion = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
