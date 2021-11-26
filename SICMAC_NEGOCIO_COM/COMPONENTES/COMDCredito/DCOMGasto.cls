VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaFiltroAplicadoCuenta(ByVal psTipoFiltro As String, ByVal bOtroFiltro As Boolean, _
Optional ByVal psOtraInstitucion = "") As ADODB.Recordset
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta

    sSql = "Select CF.nPrdConceptoCod, CF.nProdCod, "
    sSql = sSql & " CF.cAgeCod,CF.cTpoFiltro,CF.cIntitucion "
    sSql = sSql & " From ProductoConceptoFiltro CF "
    sSql = sSql & " Where CF.cTpoFiltro = '" & psTipoFiltro & "' "
    If bOtroFiltro = True Then
        sSql = sSql & " AND ISNULL(cIntitucion,'')=''"
    End If
   ' sSQL = sSQL & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    
    If Len(psOtraInstitucion) > 0 Then
        'adiciono el gasto solo de esa institucion
    sSql = sSql & " UNION"
    sSql = sSql & " Select CF.nPrdConceptoCod, CF.nProdCod, "
    sSql = sSql & " CF.cAgeCod,CF.cTpoFiltro,CF.cIntitucion "
    sSql = sSql & " From ProductoConceptoFiltro CF "
    sSql = sSql & " Where CF.cTpoFiltro = '" & psTipoFiltro & "' "
    sSql = sSql & " AND cIntitucion='" & psOtraInstitucion & "'"
    'sSQL = sSQL & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    End If
    sSql = sSql & " Order by CF.nPrdConceptoCod, CF.nProdCod, CF.cAgeCod "
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set RecuperaFiltroAplicadoCuenta = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Set oConn = Nothing

End Function


Public Function ObtenerCodInstitucionByCredito(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select cPersCod From ColocacConvenio Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerCodInstitucionByCredito = rs!cPersCod
    End If
    Set rs = Nothing
End Function


Public Function RecuperaFiltros(ByVal pnConceptoCod As Long, ByVal psTpoFiltro As String) As ADODB.Recordset
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta

    sSql = " Select CF.nProdCod,  C.cConsDescripcion, "
    sSql = sSql & " CF.cAgeCod , A.cAgeDescripcion "
    'ARCV 01-02-2007
    sSql = sSql & ", cIntitucion=ISNULL(cIntitucion,''),cPersNombre=ISNULL(P.cPersNombre,'')"
    '-------------
    sSql = sSql & " from ProductoConceptoFiltro CF"
    sSql = sSql & " INNER join ProductoConcepto PC ON CF.nPrdConceptoCod = PC.nPrdConceptoCod"
    'ALPA 20100603***********************************
    sSql = sSql & " INNER Join Constante C ON C.nConsValor = CF.nProdCod AND C.nConsCod=" & IIf(psTpoFiltro = "G", "1027", "3034")
    'sSql = sSql & " INNER Join Constante C ON C.nConsValor = CF.nProdCod AND C.nConsCod=" & IIf(psTpoFiltro = "G", "1027", "1001")
    '************************************************
    sSql = sSql & " INNER Join Agencias A ON CF.cAgeCod = A.cAgeCod"
    'ARCV 01-02-2007
    sSql = sSql & " LEFT Join Persona P ON P.cPersCod = cIntitucion "
    '--------------
    sSql = sSql & " Where CF.nPrdConceptoCod = " & pnConceptoCod & " AND cTpoFiltro = '" & psTpoFiltro & "' "
    sSql = sSql & " Order by CF.nProdCod, CF.cAgeCod"
    
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set RecuperaFiltros = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Set oConn = Nothing
End Function

Public Sub ActualizaGastosProdFiltro(ByVal pMatDatos As Variant, ByVal psTipoFiltro As String)
Dim oCred As COMDCredito.DCOMCredActBD
Dim i As Integer
Dim sSql As String
    Set oCred = New COMDCredito.DCOMCredActBD
    oCred.dBeginTrans
    
    For i = 0 To UBound(pMatDatos) - 1
        Call oCred.dInsertProductoConceptoFitro(CLng(pMatDatos(i, 0)), _
                        CInt(pMatDatos(i, 1)), pMatDatos(i, 5), psTipoFiltro, False, CStr(pMatDatos(i, 7)))
    Next i
    oCred.dCommitTrans
    Set oCred = Nothing
End Sub

Public Function RecuperaAgencias() As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim sSql As String

    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    sSql = "select cAgeCod, cAgeDescripcion From Agencias Order by cAgeCod "
    Set RecuperaAgencias = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Set oConn = Nothing
End Function

'DAOR 20061213
Public Function RecuperaGastosCabecera(ByVal pnGastosProd As GastosProd) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
'COMENTADO POR APRI20180205 ERS028-2017
'Dim lsCondicionGastos As String
'
'    If pnGastosProd = gGastosProdCredito Then
'        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '12%'   "
'    ElseIf pnGastosProd = gGastosProdJudicial Then
'        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '32%'   "
'    End If
'
'    On Error GoTo ErrorRecuperaGastosCabecera
'
'    sSql = "Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor, "
'    sSql = sSql & " nValor ,nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc, "
'    sSql = sSql & " cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual, "
'    sSql = sSql & " bAplTipCamb,bAplValorDosTit,nValorDosTit,bAplNumConCer,bAplNumMeses, bAplNumConMic "
'    sSql = sSql & " From ProductoConcepto "
'    sSql = sSql & " Where nColocCred = " & pnGastosProd & " AND nPrdConceptoCod <> 1299 " & lsCondicionGastos
'    sSql = sSql & " Order by convert(varchar(25),nPrdConceptoCod) "

    sSql = " exec stp_sel_RecuperaGastosCabecera " & pnGastosProd 'APRI20180205 ERS028-2017
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGastosCabecera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosCabecera:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description

End Function

'Public Function RecuperaGastosCabecera(ByVal pnGastosProd As GastosProd) As ADODB.Recordset
'
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim lsCondicionGastos As String
'    If pnGastosProd = gGastosProdCredito Then
'        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '12%'   "
'    ElseIf pnGastosProd = gGastosProdJudicial Then
'        lsCondicionGastos = " AND convert(varchar(15),nPrdConceptoCod) like '32%'   "
'    End If
'
'    On Error GoTo ErrorRecuperaGastosCabecera
'
'    sSql = "Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor, "
'    sSql = sSql & " nValor ,nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc, "
'    sSql = sSql & " cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual "
'    sSql = sSql & " From ProductoConcepto "
'    sSql = sSql & " Where nColocCred = " & pnGastosProd & " AND nPrdConceptoCod <> 1299 " & lsCondicionGastos
'    sSql = sSql & " Order by convert(varchar(25),nPrdConceptoCod) "
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaGastosCabecera = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorRecuperaGastosCabecera:
'    Err.Raise Err.Number, "Recupera Gastos", Err.Description
'
'End Function


Public Function ExisteGasto(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
    ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, _
    ByVal pnConcepto As Long) As Boolean
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim R As ADODB.Recordset
    
    sSql = "Select * From ColocCalendDet Where cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnColocCalendApl
    sSql = sSql & " AND nCuota = " & pnCuota
    sSql = sSql & " AND nPrdConceptoCod = " & pnConcepto
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
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

'DAOR 20061213, se consideraron nuevos criterios
Public Function RecuperaGastosAplicablesDesembolso(ByVal pnMoneda As Integer, _
    Optional ByVal psFiltro As String = "", Optional ByVal psTipoProceso As String = "'DE'", _
    Optional ByVal psGastoFijoVar As String = "N") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaGastosAplicablesDesembolso
    
    If psTipoProceso = "DE" Then psTipoProceso = "'DE'"
    
    If pnMoneda = 1 Then
        sSql = "Select * from ProductoConcepto where nInicial >0 AND "
        sSql = sSql & " nAplicado in (" & gColocConceptoAplDesembolso & "," & gColocConceptoAplTodosD & "," & gColocConceptoAplTodosDC & ") "
        sSql = sSql & " AND nPrdConceptoCod like '12%' "
        sSql = sSql & " AND ( nDeCierre = 0 OR nDeCierre is NULL) AND nColocCred = 1 AND cAplicaProceso in (" & psTipoProceso & ") "
        sSql = sSql & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
    Else
        sSql = "select  nPrdConceptoCod,cDescripcion,nAplicado,nInicial,nFinal,nTpoValor, "
        sSql = sSql & " nValor=(case nTpoValor when 1 then (case when bAplTipCamb=1 then (select nValor/nValFijo from TipoCambio "
        sSql = sSql & " where dFecCamb=(select max(dFecCamb)from TipoCambio)) else nValor  end) else nValor end), "
        sSql = sSql & " nDeCierre,nColocCred,cProducto, "
        sSql = sSql & " nMontoMin=(case when bAplTipCamb=1 then (select nMontoMin/nValFijo from TipoCambio "
        sSql = sSql & " where dFecCamb=(select max(dFecCamb)from TipoCambio)) else nMontoMin  end), "
        sSql = sSql & " nMontoMax=(case when bAplTipCamb=1 then (select nMontoMax/nValFijo from TipoCambio "
        sSql = sSql & " where dFecCamb=(select max(dFecCamb)from TipoCambio)) else nMontoMax  end), "
        sSql = sSql & " cAplicaMonto,cFiltro,cAplicaProceso, "
        sSql = sSql & " nOperador,nOperPorc,cOperMonto,nEdad,nEdadOper,nDiasApl,cGastoFijoVar,nOperDiasVenc,nDiasVenc, "
        sSql = sSql & " bSeguro,nMontoMensual,bAplTipCamb,bAplValorDosTit, "
        sSql = sSql & " nValorDosTit=(case nTpoValor when 1 then (case when bAplTipCamb=1 then (select nValorDosTit/nValFijo from TipoCambio "
        sSql = sSql & " where dFecCamb=(select max(dFecCamb)from TipoCambio)) else  nValorDosTit end)  else nValorDosTit end), "
        sSql = sSql & " bAplNumConCer,bAplNumMeses "
        sSql = sSql & " from ProductoConcepto where nInicial >0 AND "
        sSql = sSql & " nAplicado in (" & gColocConceptoAplDesembolso & "," & gColocConceptoAplTodosD & "," & gColocConceptoAplTodosDC & ") "
        sSql = sSql & " AND nPrdConceptoCod like '12%'  "
        sSql = sSql & " AND ( nDeCierre = 0 OR nDeCierre is NULL) AND nColocCred = 1 AND cAplicaProceso in (" & psTipoProceso & ") "
        sSql = sSql & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
    End If
    If pnMoneda <> -1 Then
        If pnMoneda = 1 Then
            sSql = sSql & " AND nMoneda = " & pnMoneda
        Else
            sSql = sSql & " AND (nMoneda = " & pnMoneda & " OR bAplTipCamb=1 ) "
        End If
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesDesembolso:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function

'Public Function RecuperaGastosAplicablesDesembolso(ByVal pnMoneda As Integer, _
'    Optional ByVal psFiltro As String = "", Optional ByVal psTipoProceso As String = "'DE'", _
'    Optional ByVal psGastoFijoVar As String = "N") As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'    On Error GoTo ErrorRecuperaGastosAplicablesDesembolso
'
'    If psTipoProceso = "DE" Then psTipoProceso = "'DE'"
'
'    sSql = "Select * from ProductoConcepto where nInicial >0 AND "
'    sSql = sSql & " nAplicado in (" & gColocConceptoAplDesembolso & "," & gColocConceptoAplTodosD & "," & gColocConceptoAplTodosDC & ") "
'    'sSql = sSql & " AND nPrdConceptoCod like '12%' AND (cFiltro = 'N' or cFiltro = '" & psFiltro & "') AND nMoneda = " & pnMoneda
'    sSql = sSql & " AND nPrdConceptoCod like '12%' AND nMoneda = " & pnMoneda
'    sSql = sSql & " AND ( nDeCierre = 0 OR nDeCierre is NULL) AND nColocCred = 1 AND cAplicaProceso in (" & psTipoProceso & ") "
'    sSql = sSql & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
'
'    If pnMoneda <> -1 Then
'        sSql = sSql & " AND nMoneda = " & pnMoneda
'    End If
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaGastosAplicablesDesembolso = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorRecuperaGastosAplicablesDesembolso:
'    Err.Raise Err.Number, "Recupera Gastos", Err.Description
'
'End Function


'DAOR 20061213, se consideraron nuevos criterios
Public Function RecuperaGastosAplicablesCuotas(Optional ByVal pnMoneda As Integer = -1, _
    Optional ByVal psFiltroProceso As String = "", Optional ByVal psGastoFijoVar As String = "F", _
    Optional ByVal pbGastosCierre As Boolean = False, Optional ByVal MatGastosExon As Variant = "", _
    Optional pbPolizaInc As Boolean = False, Optional pbSegMYPE As Boolean = False) As ADODB.Recordset
    'WIOR 20150709 AGREGO pbPolizaInc
    'APRI20180827 ERS061-2018 ADD pbSegMYPE
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim pMatGastosExon As String 'APRI20181025 MEJORA ERS061-2018
    
    If IsArray(MatGastosExon) Then
        If UBound(MatGastosExon) > 0 Then
            'sSql = sSql & " AND nPrdConceptoCod Not In ("
           For i = 0 To UBound(MatGastosExon) - 1
                'sSql = sSql & MatGastosExon(i) & ","
                pMatGastosExon = pMatGastosExon & MatGastosExon(i) & ","
           Next i
           pMatGastosExon = Mid(pMatGastosExon, 1, Len(pMatGastosExon) - 1)
           'sSql = Mid(sSql, 1, Len(sSql) - 1)
           'sSql = sSql & ") "
        End If
    End If
    
    sSql = "exec stp_sel_RecuperaGastosAplicablesCuotas '" & psFiltroProceso & "'," & pnMoneda & "," & IIf(pbPolizaInc, 1, 0) & "," & IIf(pbGastosCierre, 1, 0) & ",'" & psGastoFijoVar & "','" & pMatGastosExon & "'," & IIf(pbSegMYPE, 1, 0) & "" 'APRI20181025 ERS061-2018
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesCuotas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesCuotas:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function

'Public Function RecuperaGastosAplicablesCuotas(Optional ByVal pnMoneda As Integer = -1, _
'    Optional ByVal psFiltroProceso As String = "", Optional ByVal psGastoFijoVar As String = "F", _
'    Optional ByVal pbGastosCierre As Boolean = False, Optional ByVal MatGastosExon As Variant = "") As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim i As Integer
'
'    If Mid(psFiltroProceso, 1, 1) <> "'" Then
'        psFiltroProceso = "'" & psFiltroProceso & "'"
'    End If
'    On Error GoTo ErrorRecuperaGastosAplicablesCuotas
'    sSql = "Select * from ProductoConcepto "
'    'ARCV 21-07
'    'sSql = sSql & " where nValor <> 0 AND nFinal >0 AND nAplicado in (" & gColocConceptoAplCuota & "," & gColocConceptoAplTodosC & "," & gColocConceptoAplTodosDC & ") "
'    sSql = sSql & " where nInicial >0 AND nFinal >0 AND nAplicado in (" & gColocConceptoAplCuota & "," & gColocConceptoAplTodosC & "," & gColocConceptoAplTodosDC & ") "
'    '-------------
'    If pbGastosCierre Then
'        sSql = sSql & " AND (nDeCierre = 1 or nDeCierre is null) "
'    Else
'        sSql = sSql & " AND (nDeCierre = 0 or nDeCierre is null) "
'    End If
'    sSql = sSql & " AND nColocCred = 1 AND nPrdConceptoCod like '12%' "
'    If psFiltroProceso <> "" Then
'        sSql = sSql & " AND cAplicaProceso in (" & IIf(psFiltroProceso = "'CA'" Or psFiltroProceso = "'PP'" Or psFiltroProceso = "'PA'", "'PA','CA','PP'", psFiltroProceso) & ")"
'    End If
'    sSql = sSql & " AND ( cGastoFijoVar  = 'N' or cGastoFijoVar = '" & psGastoFijoVar & "')  "
'
'    If pnMoneda <> -1 Then
'        sSql = sSql & " AND nMoneda = " & pnMoneda
'    End If
'
'    If IsArray(MatGastosExon) Then
'        If UBound(MatGastosExon) > 0 Then
'            sSql = sSql & " AND nPrdConceptoCod Not In ("
'           For i = 0 To UBound(MatGastosExon) - 1
'                sSql = sSql & MatGastosExon(i) & ","
'           Next i
'           sSql = Mid(sSql, 1, Len(sSql) - 1)
'           sSql = sSql & ") "
'        End If
'    End If
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaGastosAplicablesCuotas = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorRecuperaGastosAplicablesCuotas:
'    Err.Raise Err.Number, "Recupera Gastos", Err.Description
'
'End Function


Public Function UltimoCorrelXProd() As Long
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    sSql = "Select MAX(convert(varchar(25),nPrdConceptoCod)) as nUlt From ProductoConcepto where nPrdConceptoCod like '12%' and nPrdConceptoCod <> 1299 "
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    UltimoCorrelXProd = IIf(IsNull(R!nUlt), 0, R!nUlt)
    UltimoCorrelXProd = UltimoCorrelXProd
    
End Function

'DAOR 20061213, Se consideraron nuevos campos :
'pnAplTipCamb : Especifica si se aplica el tipo de cambio para cuentas en dolares
'pnAplValorDosTit : Especifica si se aplica la restricción de nuevo valor en caso que el credito tenga mas de un titular
'pnValorDosTit: Monto del valor aplicado para mas de dos titulares
Public Sub NuevoGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
    psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
    ByVal pnDiasApl As Integer, ByVal pnGastoTipoProd As Integer, ByVal pnEdadOper As Integer, _
    ByVal psGastoFijoVar As String, ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer, _
    ByVal psMontoMens As String, pnAplTipCamb As Integer, pnAplValorDosTit As Integer, pnValorDosTit As Double, _
    pnAplNumConCer As Integer, pnAplNumMeses As Integer, Optional pnAplNumComMic As Integer, _
    Optional pnPagoConvenio As Integer = 0)
    
    'CTI2 20181226 ADD pbPagoConvenio, ERS075-2018
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorNuevoGasto
    sSql = "INSERT INTO ProductoConcepto(nPrdConceptoCod, cDescripcion, nAplicado, nInicial, "
    sSql = sSql & " nFinal, nTpoValor, nValor, nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,"
    sSql = sSql & " cAplicaProceso, nOperador, nOperPorc,cOperMonto,nEdad,nDiasApl,nColocCred,nEdadOper,"
    sSql = sSql & " cGastoFijoVar,nOperDiasVenc,nDiasVenc, bSeguro,nMontoMensual,bAplTipCamb,bAplValorDosTit,"
    sSql = sSql & " nValorDosTit,bAplNumConCer,bAplNumMeses,bAplNumConMic, nAplicaConvenio)"
    sSql = sSql & " VALUES(" & psColocConceptoCod & ",'" & psDescripcion & "'," & psAplicado & "," & Format(pnInicial, "#0.00") & "," & Format(pnFinal, "#0.00") & "," & psTpoValor & "," & Format(pnValor, "#0.0000") & "," & psMoneda & "," & Format(pnMontoMin, "#0.00") & "," & Format(pnMontoMax, "#0.00") & ",'" & psAplMonto & "','" & psAplGarTpo & "',"
    sSql = sSql & "'" & psAplicaproceso & "'," & IIf(pnOperador = -1, "NULL", pnOperador) & "," & IIf(pnOperador = -1, "NULL", Format(pnOperPorc, "#0.00")) & "," & IIf(pnOperador = -1, "NULL", "'" & psOperMonto & "'") & "," & pnEdad & "," & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & "," & pnGastoTipoProd & "," & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ","
    sSql = sSql & "'" & psGastoFijoVar & "'," & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & "," & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & "," & pnSeguro & "," & psMontoMens & "," & pnAplTipCamb & "," & pnAplValorDosTit & "," & IIf(pnValorDosTit = -1, "NULL", Format(pnValorDosTit, "#0.0000")) & ","
    sSql = sSql & pnAplNumConCer & "," & pnAplNumMeses & "," & pnAplNumComMic & ", " & pnPagoConvenio & ")"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoGasto:
    Err.Raise Err.Number, "Nuevo Gasto", Err.Description
End Sub

'Public Sub NuevoGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
'    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
'    psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
'    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
'    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
'    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
'    ByVal pnDiasApl As Integer, ByVal pnGastoTipoProd As Integer, ByVal pnEdadOper As Integer, _
'    ByVal psGastoFijoVar As String, ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer, _
'    ByVal psMontoMens As String)
'
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'    On Error GoTo ErrorNuevoGasto
'    sSql = "INSERT INTO ProductoConcepto(nPrdConceptoCod, cDescripcion, nAplicado, nInicial, "
'    sSql = sSql & " nFinal, nTpoValor, nValor, nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,"
'    sSql = sSql & " cAplicaProceso, nOperador, nOperPorc,cOperMonto,nEdad,nDiasApl,nColocCred,nEdadOper,cGastoFijoVar,nOperDiasVenc,nDiasVenc, bSeguro,nMontoMensual)"
'    sSql = sSql & " VALUES(" & psColocConceptoCod & ",'" & psDescripcion & "'," & psAplicado & "," & Format(pnInicial, "#0.00") & "," & Format(pnFinal, "#0.00") & "," & psTpoValor & "," & Format(pnValor, "#0.0000") & "," & psMoneda & "," & Format(pnMontoMin, "#0.00") & "," & Format(pnMontoMax, "#0.00") & ",'" & psAplMonto & "','" & psAplGarTpo & "',"
'    sSql = sSql & "'" & psAplicaproceso & "'," & IIf(pnOperador = -1, "NULL", pnOperador) & "," & IIf(pnOperador = -1, "NULL", Format(pnOperPorc, "#0.00")) & "," & IIf(pnOperador = -1, "NULL", "'" & psOperMonto & "'") & "," & pnEdad & "," & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & "," & pnGastoTipoProd & "," & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ",'" & psGastoFijoVar & "'," & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & "," & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & "," & pnSeguro & ",'" & psMontoMens & "')"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorNuevoGasto:
'    Err.Raise Err.Number, "Nuevo Gasto", Err.Description
'End Sub


'DAOR 20061213, Se consideraron nuevos campos :
'pnAplTipCamb : Especifica si se aplica el tipo de cambio para cuentas en dolares
'pnAplValorDosTit : Especifica si se aplica la restricción de nuevo valor en caso que el credito tenga mas de un titular
'pnValorDosTit: Monto del valor aplicado para mas de dos titulares
Public Sub ActualizarGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
    ByVal psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
    ByVal pnDiasApl As Integer, ByVal pnEdadOper As Integer, ByVal psGastoFijoVar As String, _
    ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer, _
    ByVal psMontoMens As String, pnAplTipCamb As Integer, pnAplValorDosTit As Integer, pnValorDosTit As Double, _
    pnAplNumConCer As Integer, pnAplNumMeses As Integer, pnAplNumConMic As Integer, _
    Optional ByVal pnPagoConvenio As Integer = 0)
    
    'CTI2 20181226 ADD pbPagoConvenio, ERS075-2018
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarGasto
    sSql = " UPDATE ProductoConcepto SET "
    sSql = sSql & "cDescripcion = '" & psDescripcion & "',"
    sSql = sSql & "nAplicado = " & psAplicado & ","
    sSql = sSql & "nInicial = " & Format(pnInicial, "#0.00") & ","
    sSql = sSql & "nFinal = " & Format(pnFinal, "#0.00") & ","
    sSql = sSql & "nTpoValor = " & psTpoValor & ","
    sSql = sSql & "nValor = " & Format(pnValor, "#0.0000") & ","
    sSql = sSql & "nMoneda = " & psMoneda & ", "
    sSql = sSql & " nMontoMin =  " & Format(pnMontoMin, "#0.00") & ","
    sSql = sSql & " nMontoMax =  " & Format(pnMontoMax, "#0.00") & ","
    sSql = sSql & " cAplicaMonto =  '" & psAplMonto & "',"
    sSql = sSql & " cFiltro = '" & psAplGarTpo & "', "
    sSql = sSql & " cAplicaProceso = '" & psAplicaproceso & "',"
    sSql = sSql & " nOperador = " & IIf(pnOperador = -1, "NULL", pnOperador) & ","
    sSql = sSql & " nOperPorc = " & IIf(pnOperPorc = -1, "NULL", Format(pnOperPorc, "#0.00")) & ","
    sSql = sSql & " cOperMonto = '" & psOperMonto & "',"
    sSql = sSql & " nEdad = " & pnEdad & ","
    sSql = sSql & " nEdadOper = " & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ","
    sSql = sSql & " nDiasApl = " & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & ","
    sSql = sSql & " cGastoFijoVar = '" & psGastoFijoVar & "', "
    sSql = sSql & " nOperDiasVenc = " & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & ", "
    sSql = sSql & " nDiasVenc = " & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & ","
    sSql = sSql & " bSeguro = " & pnSeguro & ","
    sSql = sSql & " nMontoMensual=" & psMontoMens & ","
    sSql = sSql & " bAplTipCamb=" & pnAplTipCamb & ","
    sSql = sSql & " bAplValorDosTit=" & pnAplValorDosTit & ","
    sSql = sSql & " nValorDosTit=" & IIf(pnValorDosTit = -1, "NULL", Format(pnValorDosTit, "#0.0000")) & ","
    sSql = sSql & " bAplNumConCer=" & pnAplNumConCer & ","
    sSql = sSql & " bAplNumMeses=" & pnAplNumMeses & ","
    sSql = sSql & " bAplNumConMic=" & pnAplNumConMic & ","
    sSql = sSql & " nAplicaConvenio=" & pnPagoConvenio & " "
    sSql = sSql & " WHERE nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizarGasto:
    Err.Raise Err.Number, "Actualizar Gasto", Err.Description
End Sub

'Public Sub ActualizarGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, _
'    ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, _
'    ByVal psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String, _
'    ByVal pnMontoMin As Double, ByVal pnMontoMax As Double, ByVal psAplMonto As String, _
'    ByVal psAplGarTpo As String, ByVal psAplicaproceso As String, ByVal pnOperador As Integer, _
'    ByVal pnOperPorc As Double, ByVal psOperMonto As String, ByVal pnEdad As Integer, _
'    ByVal pnDiasApl As Integer, ByVal pnEdadOper As Integer, ByVal psGastoFijoVar As String, _
'    ByVal pnOperDiasVenc As Integer, ByVal pnDiasVenc As Integer, ByVal pnSeguro As Integer, _
'    ByVal psMontoMens As String)
'
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'
'
'    On Error GoTo ErrorActualizarGasto
'    sSql = " UPDATE ProductoConcepto SET "
'    sSql = sSql & "cDescripcion = '" & psDescripcion & "',"
'    sSql = sSql & "nAplicado = " & psAplicado & ","
'    sSql = sSql & "nInicial = " & Format(pnInicial, "#0.00") & ","
'    sSql = sSql & "nFinal = " & Format(pnFinal, "#0.00") & ","
'    sSql = sSql & "nTpoValor = " & psTpoValor & ","
'    sSql = sSql & "nValor = " & Format(pnValor, "#0.0000") & ","
'    sSql = sSql & "nMoneda = " & psMoneda & ", "
'    sSql = sSql & " nMontoMin =  " & Format(pnMontoMin, "#0.00") & ","
'    sSql = sSql & " nMontoMax =  " & Format(pnMontoMax, "#0.00") & ","
'    sSql = sSql & " cAplicaMonto =  '" & psAplMonto & "',"
'    sSql = sSql & " cFiltro = '" & psAplGarTpo & "', "
'    sSql = sSql & " cAplicaProceso = '" & psAplicaproceso & "',"
'    sSql = sSql & " nOperador = " & IIf(pnOperador = -1, "NULL", pnOperador) & ","
'    sSql = sSql & " nOperPorc = " & IIf(pnOperPorc = -1, "NULL", Format(pnOperPorc, "#0.00")) & ","
'    sSql = sSql & " cOperMonto = '" & psOperMonto & "',"
'    sSql = sSql & " nEdad = " & pnEdad & ","
'    sSql = sSql & " nEdadOper = " & IIf(pnEdadOper = -1, "NULL", pnEdadOper) & ","
'    sSql = sSql & " nDiasApl = " & IIf(pnDiasApl = -1, "NULL", pnDiasApl) & ","
'    sSql = sSql & " cGastoFijoVar = '" & psGastoFijoVar & "', "
'    sSql = sSql & " nOperDiasVenc = " & IIf(pnOperDiasVenc = -1, "NULL", pnOperDiasVenc) & ", "
'    sSql = sSql & " nDiasVenc = " & IIf(pnDiasVenc = -1, "NULL", pnDiasVenc) & ","
'    sSql = sSql & " bSeguro = " & pnSeguro & ","
'    sSql = sSql & " nMontoMensual=" & psMontoMens
'    sSql = sSql & " WHERE nPrdConceptoCod = " & psColocConceptoCod
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorActualizarGasto:
'    Err.Raise Err.Number, "Actualizar Gasto", Err.Description
'End Sub

Public Function GastoAsignadoACredito(ByVal psColocConceptoCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    
    sSql = "Select * From ColocCalendDet Where nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        GastoAsignadoACredito = True
        R.Close
        Exit Function
        oConecta.CierraConexion
    Else
        GastoAsignadoACredito = False
    End If
    R.Close
    
    sSql = "Select * From MovColDet Where nPrdConceptoCod = " & psColocConceptoCod
    Set R = oConecta.CargaRecordSet(sSql)
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
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorEliminarGasto
    sSql = " DELETE ProductoConcepto "
    sSql = sSql & " WHERE nPrdConceptoCod = " & psColocConceptoCod
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarGasto:
    Err.Raise Err.Number, "Eliminar Gasto", Err.Description
End Sub

Public Function RecuperaGastoValorPenalidadCancelacion(ByVal pnMoneda As Integer) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    On Error GoTo ErrorRecuperaGastoValorPenalidadCancelacion
    If pnMoneda = gMonedaNacional Then
        sSql = "Select nValor from ProductoConcepto where nPrdConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoSoles
    Else
        sSql = "Select nValor from ProductoConcepto where nPrdConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoDolares
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
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
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select cPersNombre From Persona Where cPersCod='" & psCodInstitucion & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        RecupNomInstitucion = rs!cPersNombre
    End If
    Set rs = Nothing
End Function


Public Function RecupCodInstitucion(ByVal psCodProductoConcepto As String, ByVal sProducto As String, ByVal psAgenciaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select Pers.cPersCod,Pers.cPersNombre"
    sSql = sSql & " From PRODUCTOCONCEPTOFILTRO PCF"
    sSql = sSql & " Inner Join  Persona Pers on Pers.cPersCod=PCF.cIntitucion"
    sSql = sSql & " Where nPrdConceptoCod=" & psCodProductoConcepto & " and nProdCod=" & sProducto & " and cAgeCod='" & psAgenciaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecupCodInstitucion = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function CargarControlesGastos(ByRef prsApli As ADODB.Recordset, _
                                     ByRef prsProd As ADODB.Recordset, _
                                     ByRef prsOpe As ADODB.Recordset, _
                                     ByRef prsEdad As ADODB.Recordset, _
                                     ByRef prsDias As ADODB.Recordset) As Recordset
    
    Dim oCons As COMDConstantes.DCOMConstantes
    
    
    On Error GoTo ErrorCargarControlesGastos
        
    Set oCons = New COMDConstantes.DCOMConstantes
    
    Set prsApli = oCons.RecuperaConstantes(gColocCalendApl, gCredFiltroCalendApl)
    Set prsProd = oCons.RecuperaConstantes(gGastosProd)
    Set prsOpe = oCons.RecuperaConstantes(6031, gCredFiltroGastosOperad)
    Set prsEdad = oCons.RecuperaConstantes(6031, gCredFiltroGastosOperad)
    Set prsDias = oCons.RecuperaConstantes(6031, gCredFiltroGastosOperad)
    
    Set oCons = Nothing
    Exit Function
    
ErrorCargarControlesGastos:
    Err.Raise Err.Number, "Cargar Controles Gastos", Err.Description
    
End Function

Public Function CargarDatosPenalidad(ByVal psCtaCod As String, _
                                    ByRef prsDatos As ADODB.Recordset, _
                                     ByRef prsGastos As ADODB.Recordset, _
                                     ByRef prsGastosExo As ADODB.Recordset) As Recordset
    
Dim oCredito As COMDCredito.DCOMCredito

    On Error GoTo ErrorCargarDatosPenalidad
    
    Set oCredito = New COMDCredito.DCOMCredito
    Set prsDatos = oCredito.RecuperaDatosComunes(psCtaCod, False)
    Set prsGastosExo = oCredito.RecuperaGastosExonerados(psCtaCod)
    Set oCredito = Nothing
    
    Set prsGastos = RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), "PA")
    
    Exit Function
    
            
ErrorCargarDatosPenalidad:
    Err.Raise Err.Number, "Cargar Datos Penalidad", Err.Description
    
End Function

'MAVM 20120607 ***
Public Function GetGastoAdmMiViv(ByVal psConcpetoCod As String) As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select nValor From ProductoConcepto Where nPrdConceptoCod='" & psConcpetoCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        GetGastoAdmMiViv = rs!nValor
    End If
    Set rs = Nothing
End Function
'***
'RECO20150526 ERS023-2015 *********************************
Public Function RecuperaGastosMYPE(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "stp_sel_RecuperaGastosMYPE '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaGastosMYPE = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'RECO FIN *************************************************

'JUEZ 20150714 ***************************************************
Public Function RecuperaProductoConceptoHist(ByVal pnConceptoCod As Integer, ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "stp_sel_RecuperaProductoConceptoHist " & pnConceptoCod & ",'" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaProductoConceptoHist = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END JUEZ ********************************************************

