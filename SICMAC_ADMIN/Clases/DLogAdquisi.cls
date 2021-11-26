VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogAdquisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Public Enum TpoCargaSel
    SelUnRegistro = 1
    SelTodosGnral = 5
    SelTodosEstado = 6
    SelTodosEstadoParaCotizar = 7
    SelTodosEstadoParaDesierto = 8
    SelTodosEstadoParaAceptar = 9
    SelTodosEstadoParaCancelar = 10
    SelTodosEstadoParaProAdj = 11
    SelTodosEstadoParaConsenti = 12
End Enum

Public Enum TpoCargaSelCot
    SelCotPersona = 1
    SelCotNroCotPersona = 2
    SelCotGarantia = 3
End Enum

Public Enum TpoCargaSelCotDet
    SelCotDetUnRegistro = 1
    SelCotDetUnRegistroCotiza = 2
    SelCotDetUnRegEvalua = 5
    SelCotDetUnRegEvaluaTodos = 6
End Enum

Public Enum TpoCargaSelCotPar
    SelCotParRegistro = 1
    SelCotParRegEvalua = 5
End Enum
Public Enum TpoLogSelTipoBs
    SelBienesSuministros = 1
    SelServiciosConsultorias = 2
    SelObras = 3
End Enum

Public Enum TpoLogSelEvaluacion
    SelEvalTecnica = 1
    SelEvalEconomica = 2
End Enum

Public Enum TpoLogSelTipoSeleccion
    SelMenorCuantia = 1
    SelDirectaSelectiva = 2
    SelDirectaPublica = 3
    SelLicitaciónPublica = 4
End Enum
Public Enum TpoLogSelEstProveedor
    SelProCalificado = 0
    SelProDesCalificado = 1
End Enum
Public Enum TpoLogSelEstProceso
    SelEstProcesoIniciado = 0
    SelEstProcesoEvaluacionTec = 1
    SelEstProcesoEvaluacionEco = 2
    SelEstProcesoFinEvaluacion = 3
    SelEstProcesoCerrado = 4
    SelEstProcesoCancelado = 5
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
        vsConexion = oIni.CadenaConexion(App.Path & "\SicmacIAdmin.INI")
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

'Devuelve un RecordSet con LogSelDetalle
Public Function CargaSelDetalle(ByVal pnSelNro As Long, Tipo As Integer, Optional ByVal psCodProvedor As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        Select Case Tipo
        Case 1
            sSql = "Select  s.cBSCod,bs.cBSDescripcion ,  " _
            & "  t.cConsDescripcion , nBSUnidad,s.cDescripcionBien, s.nLogSelDetCantidad, s.nLogSelDetPrecio,s.nLogSelDetCantidad * s.nLogSelDetPrecio as Monto" _
            & "  from logseldetalle s,bienesServicios bs,Constante t " _
            & "  Where bs.cBSCod = s.cBSCod " _
            & "  and bs.nBSUnidad = t.nConsValor and t.nConsCod = " & gUnidadMedida & " and nLogSelNro = " & pnSelNro & " "
        Case 3 'Cotizacion
            sSql = " Select s.cBSCod,bs.cBSDescripcion,t.cConsDescripcion , " _
            & " nBSUnidad,d.cDescripcionBien,s.nLogSelCotDetCantidad,s.nLogSelCotDetPrecio, " _
            & " s.nLogSelCotDetCantidad * s.nLogSelCotDetPrecio as Monto,indicadorParticipa " _
            & " from LogSelCotDetalle s,bienesServicios bs,Constante t,logseldetalle d " _
            & " Where bs.cBSCod = s.cBSCod   and " _
            & " d.nLogSelNro = s.nLogSelNro and d.cBSCod = s.cBSCod and " _
            & " bs.nBSUnidad = t.nConsValor and " _
            & " t.nConsCod = " & gUnidadMedida & " And s.nLogSelNro = " & pnSelNro & " and cPersCod ='" & psCodProvedor & "'  "
        Case 4
            sSql = " Select s.cBSCod,bs.cBSDescripcion,t.cConsDescripcion , " _
            & " nBSUnidad,d.cDescripcionBien,s.nLogSelCotDetCantidad,s.nLogSelCotDetPrecio, " _
            & " s.nLogSelCotDetCantidad * s.nLogSelCotDetPrecio as Monto,indicadorParticipa " _
            & " from LogSelCotDetalle s,bienesServicios bs,Constante t,logseldetalle d " _
            & " Where bs.cBSCod = s.cBSCod   and " _
            & " d.nLogSelNro = s.nLogSelNro and d.cBSCod = s.cBSCod and " _
            & " bs.nBSUnidad = t.nConsValor and " _
            & " t.nConsCod = " & gUnidadMedida & " And s.nLogSelNro = " & pnSelNro & " and cPersCod ='" & psCodProvedor & "'  " _
            & " and  nIndicaGanador  = 'SI'"
        End Select
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelDetalle = rs
End Function


'Devuelve un RecordSet con LogSeleccion
Public Function CargaSeleccion(ByVal psTpoCarga As TpoCargaSel, Optional ByVal pnSelNro As Long, _
Optional ByVal psEstUno As LogSelEstado, Optional ByVal psEstDos As LogSelEstado, _
Optional ByVal psEstTre As LogSelEstado, Optional ByVal psEstCua As LogSelEstado) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = SelUnRegistro Then
            sSql = "SELECT s.nLogSelNro, s.dLogSelRes, s.cLogSelResNro, s.nLogSelEstado, " & _
                "       s.cAreaCod, a.cAreaDescripcion, s.nLogSelCostoBase, s.nLogSelValorRefe, " & _
                "       IsNull(s.nLogSelSisAdj,0) nLogSelSisAdj, IsNull(ca.cConsDescripcion,'') cConsDescSisAdj, " & _
                "       IsNull(s.nLogSelMoneda,0) nLogSelMoneda, " & _
                "       IsNull(s.nLogSelTipCambio,0) nLogSelTipCambio,  " & _
                "       IsNull(s.nLogSelTpo,0) nLogSelTpo, IsNull(ct.cConsDescripcion,'') cConsDescSelTpo, " & _
                "       IsNull(s.nLogSelCotNro, 0) nLogSelCotNro " & _
                " FROM LogSeleccion S LEFT JOIN Areas A ON s.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante CA ON s.nLogSelSisAdj = ca.nConsValor " & _
                "           AND ca.nConsCod = " & gLogSelSisAdj & " " & _
                "       LEFT JOIN Constante CT ON s.nLogSelTpo = ct.nConsValor " & _
                "           AND ct.nConsCod = " & gLogSelTpo & "  " & _
                " WHERE s.nLogSelNro = " & pnSelNro & " "
        ElseIf psTpoCarga = SelTodosGnral Then
            sSql = "SELECT s.nLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE NOT c.nConsValor IN (" & psEstUno & "," & psEstDos & "," & psEstTre & ")"
        ElseIf psTpoCarga = SelTodosEstado Then
            sSql = "SELECT m.cMovNro Código, s.cLogSelResNro Resolución, c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogSeleccion S ON m.nMovNro = s.nLogSelNro LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado IN (" & psEstUno & "," & psEstDos & "," & psEstTre & "," & psEstCua & ")"
        ElseIf psTpoCarga = SelTodosEstadoParaCotizar Then
            sSql = "SELECT m.cMovNro Código, s.cLogSelResNro Resolución, c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogSeleccion S ON m.nMovNro = s.nLogSelNro LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado >= " & gLogSelEstadoCotizacion & " "
        ElseIf psTpoCarga = SelTodosEstadoParaDesierto Then
            sSql = "SELECT s.nLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado >= " & gLogSelEstadoCotizacion & " " & _
                "       AND NOT nLogSelEstado IN (" & gLogSelEstadoAceptado & "," & _
                        gLogSelEstadoRechazado & "," & gLogSelEstadoDesierto & ")"
        ElseIf psTpoCarga = SelTodosEstadoParaAceptar Then
            sSql = "SELECT s.nLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado = " & gLogSelEstadoProcAdju & " "
        ElseIf psTpoCarga = SelTodosEstadoParaCancelar Then
            sSql = "SELECT s.nLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE NOT nLogSelEstado IN (" & gLogSelEstadoAceptado & "," & _
                        gLogSelEstadoRechazado & "," & gLogSelEstadoDesierto & ")"
        ElseIf psTpoCarga = SelTodosEstadoParaProAdj Then
            sSql = "SELECT m.cMovNro Código, s.cLogSelResNro Resolución, c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogSeleccion S ON m.nMovNro = s.nLogSelNro LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado IN (" & gLogSelEstadoConsulta & "," & _
                        gLogSelEstadoAbsolucion & "," & gLogSelEstadoObservacion & "," & _
                        gLogSelEstadoRegBase & "," & gLogSelEstadoCotizacion & ")"
        ElseIf psTpoCarga = SelTodosEstadoParaConsenti Then
            sSql = "SELECT s.nLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.nLogSelEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogSelEstado & " " & _
                " WHERE nLogSelEstado IN (" & gLogSelEstadoAceptado & ")"
        Else
            sSql = ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSeleccion = rs
End Function

'Devuelve un RecordSet con LogSelComite
Public Function CargaSelComite(ByVal pnSelNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT sc.cPersCod, p.cPersNombre," & _
            "        sc.cAreaCod, a.cAreaDescripcion  " & _
            " FROM LogSelComite SC LEFT JOIN Areas A ON sc.cAreaCod = a.cAreaCod " & _
            "       LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
            " WHERE sc.nLogSelNro = " & pnSelNro & " "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelComite = rs
End Function

Public Function ValidaLogSelDetalle(ByVal pnSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  Count(*) as cantidad from logseldetalle where nLogSelNro = " & pnSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    ValidaLogSelDetalle = rs!Cantidad
End Function
Public Function ValidaLogSelDetalleCotizaProveedor(ByVal pnSelNro As Long, ByVal psCodPers As String) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  count(*) as cantidad from LogSelCotDetalle where nLogSelNro = " & pnSelNro & " and  cPersCod = '" & psCodPers & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    ValidaLogSelDetalleCotizaProveedor = rs!Cantidad
End Function

Public Function CuentaLogSelProveedor(ByVal pnSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select count(distinct cPersCod) as Total from  LogSelProveedor  where  nLogSelNro  ='" & pnSelNro & "'  "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaLogSelProveedor = rs!Total
End Function

Public Function CuentaLogSelProveedorEvalTec(ByVal pnSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select count(distinct cPersCod) as Total from  LogSelCalificacionTecnica  where  nLogSelNro  ='" & pnSelNro & "'  "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaLogSelProveedorEvalTec = rs!Total
End Function
Public Function CuentaLogSelProveedorEvalEco(ByVal pnSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select count(distinct cPersCod) as Total from  LogSelCotDetalle  where  nLogSelNro  ='" & pnSelNro & "'  "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaLogSelProveedorEvalEco = rs!Total
End Function



Public Function CargaSeleccionProceso(ByVal pnSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "select nLogSelNro,nPeriodo,cCodTipoBS,nLogSelNumeroCot, " _
            & " cCodTipoSelec,nLogSelEstado,nLogSelMoneda,nLogSelTipoProceso, " _
            & " Substring(cFechaConvoSeleccion,1,2)+'/'+Substring(cFechaConvoSeleccion,3,2)+'/'+Substring(cFechaConvoSeleccion,5,4) as FechaConvo ," _
            & " nRangoIni, nRangoFin ,cDescripcionProceso" _
            & " from logseleccionP  where nLogSelNro =  " & pnSelNro & " "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaSeleccionProceso = rs
End Function

Public Function CargaSeleccionComite(ByVal pnSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "Select s.cPersCod,p.cPersNombre,c.cDescripcion + replicate(' ',100) +  convert(char(1),s.nCargoCod)   from logselcomite s,persona p, logselcargos c " _
                 & " where  s.cPersCod = p.cPersCod and c.nCargoCod = s.nCargoCod and " _
                 & " nLogSelNro =  " & pnSelNro & " "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaSeleccionComite = rs
End Function
Public Function CargaSeleccionPuntajes(ByVal pnSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "select nLogSelNro , cCodTipoEvaluacion, nPuntajeMaximo, nPesoPonderado, nPuntajeMinimo " _
                   & "  From LogSelConfigPuntajes  where nLogSelNro =  " & pnSelNro & " "
            Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    'and cCodTipoEvaluacion =1
    End If
    Set CargaSeleccionPuntajes = rs
End Function

'Devuelve un RecordSet con LogSelCronograma
Public Function CargaSelCronograma(ByVal pnSelNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT sc.nLogSelCroNro, c.cConsDescripcion, 1, " & _
            "       Convert(varchar(12), sc.dLogSelCroIni, 103), " & _
            "       Convert(varchar(12), sc.dLogSelCroFin, 103)  " & _
            " FROM LogSelCronograma SC LEFT JOIN Constante C ON sc.nLogSelCroNro = c.nConsValor " & _
            "       AND c.nConsCod = " & 5054 & "" & _
            " WHERE sc.nLogSelNro = " & pnSelNro & " "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCronograma = rs
End Function

'Devuelve un RecordSet con LogSelParametro
Public Function CargaSelParametro(ByVal pnSelNro As Integer, ByVal pnSelParTpo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT sp.nLogSelParNro, cp.cConsDescripcion cConsDescParametro, " & _
            "       sp.nLogSelParPuntaje, (ct.cConsDescripcion + Space(40) +  Convert(Char(1),sp.nLogSelParNroTpo)) cConsDescNroTpo " & _
            " FROM LogSelParametro SP LEFT JOIN Constante CP ON sp.nLogSelParNro = cp.nConsValor " & _
            "           AND cp.nConsCod = " & IIf(pnSelParTpo = 1, gLogSelParTec, gLogSelParEco) & " " & _
            "       LEFT JOIN Constante CT ON sp.nLogSelParNroTpo = ct.nConsValor " & _
            "           AND ct.nConsCod = " & gLogSelParTpo & " " & _
            " WHERE sp.nLogSelNro = " & pnSelNro & " AND sp.nLogSelParTpo = " & pnSelParTpo & " " & _
            " ORDER BY sp.nLogSelParNro"
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelParametro = rs
End Function

'Devuelve un RecordSet con LogSelParDetalle
Public Function CargaSelParDetalle(ByVal pnSelNro As Integer, ByVal pnSelParTpo As Integer, _
Optional ByVal pnSelParNro As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnSelParNro = 0 Then
            sSql = "SELECT spd.nLogSelParNro, spd.nLogSelParDetNro, spd.nLogSelParDetIni, " & _
                "       spd.nLogSelParDetFin, spd.nLogSelParDetPuntaje" & _
                " FROM LogSelParDetalle SPD " & _
                " WHERE spd.nLogSelNro = " & pnSelNro & " AND spd.nLogSelParTpo = " & pnSelParTpo & " " & _
                " ORDER BY spd.nLogSelParDetNro"
        Else
            sSql = "SELECT spd.nLogSelParNro, spd.nLogSelParDetNro, spd.nLogSelParDetIni, " & _
                "       spd.nLogSelParDetFin, spd.nLogSelParDetPuntaje" & _
                " FROM LogSelParDetalle SPD " & _
                " WHERE spd.nLogSelNro = " & pnSelNro & " AND spd.nLogSelParTpo = " & pnSelParTpo & " " & _
                "       AND spd.nLogSelParNro = " & pnSelParNro & " " & _
                " ORDER BY spd.nLogSelParDetNro"
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelParDetalle = rs
End Function

'Devuelve un RecordSet con LogSelCotPar
Public Function CargaSelCotPar(ByVal psTpoCarga As TpoCargaSelCotPar, _
ByVal pnSelNro As Integer, ByVal pnSelParTpo As Integer, _
Optional ByVal pnSelCotNro As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = SelCotParRegistro Then
            If pnSelCotNro = 0 Then
                sSql = "SELECT sp.nLogSelParNro, cp.cConsDescripcion cConsDescParametro, " & _
                    "       sp.nLogSelCotParValor " & _
                    " FROM LogSelCotPar SP LEFT JOIN Constante CP ON sp.nLogSelParNro = cp.nConsValor " & _
                    "           AND cp.nConsCod = " & IIf(pnSelParTpo = 1, gLogSelParTec, gLogSelParEco) & " " & _
                    " WHERE sp.nLogSelNro = " & pnSelNro & " AND sp.nLogSelParTpo = " & pnSelParTpo & " " & _
                    " ORDER BY sp.nLogSelParNro"
            Else
                sSql = "SELECT sp.nLogSelParNro, cp.cConsDescripcion cConsDescParametro, " & _
                    "       sp.nLogSelCotParValor " & _
                    " FROM LogSelCotPar SP LEFT JOIN Constante CP ON sp.nLogSelParNro = cp.nConsValor " & _
                    "           AND cp.nConsCod = " & IIf(pnSelParTpo = 1, gLogSelParTec, gLogSelParEco) & " " & _
                    " WHERE sp.nLogSelNro = " & pnSelNro & " AND sp.nLogSelParTpo = " & pnSelParTpo & " " & _
                    "       AND sp.nLogSelCotNro = " & pnSelCotNro & " " & _
                    " ORDER BY sp.nLogSelParNro"
            End If
        ElseIf psTpoCarga = SelCotParRegEvalua Then
                'sSQL = "SELECT sp.nLogSelParNro, cp.cConsDescripcion cConsDescParametro, " & _
                    "       sp.nLogSelCotParValor " & _
                    " FROM LogSelCotPar SP LEFT JOIN Constante CP ON sp.nLogSelParNro = cp.nConsValor " & _
                    "           AND cp.nConsCod = " & IIf(pnSelParTpo = 1, gLogSelParTec, gLogSelParEco) & " " & _
                    " WHERE sp.nLogSelNro = " & pnSelNro & " AND sp.nLogSelParTpo = " & pnSelParTpo & " " & _
                    "       AND sp.nLogSelCotNro = " & pnSelCotNro & " "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCotPar = rs
End Function

'Devuelve un RecordSet con LogSelPostor
Public Function CargaSelPostor(ByVal psSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT sp.cPersCod, p.cPersNombre, Convert(varchar(12), sp.dLogSelPosEntrega,103), " & _
            "       isnull(sp.cLogSelPosConsulta,'') cLogSelPosConsulta, " & _
            "       isnull(sp.cLogSelPosAbsolucion,'') cLogSelPosAbsolucion, " & _
            "       isnull(sp.cLogSelPosObservacion,'') cLogSelPosObservacion " & _
            " FROM LogSelPostor SP LEFT JOIN " & vsCentralPer & "Persona P ON sp.cPersCod = p.cPersCod " & _
            " WHERE sp.nLogSelNro = " & psSelNro & " "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelPostor = rs
End Function

Public Function CargaSelcargos() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  cDescripcion ,nCargoCod   from LogSelcargos where   nStatusRegistro = 1"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelcargos = rs
End Function


Public Function CargaSelCriteriosTecnicos(contexto As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
       Select Case contexto
            Case 1
            sSql = "select  cCodCriterioTec as Codigo,cDescripcion as Descripcion,nPuntajeDefault as Puntaje,cUltimaActualizacion  from LogSelCriteriosTecnicos where   nStatusRegistro = 1"
            Case 2
            sSql = "Select  0 as cCodCriterioTec,'Criterios de Evaluacion Tecnica' as cDescripcion , 2 " _
                   & " Union " _
                   & " Select  cCodCriterioTec ,cDescripcion  ,3 from LogSelCriteriosTecnicos where   nStatusRegistro = 1  order by cCodCriterioTec"
       End Select
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCriteriosTecnicos = rs
End Function

Public Function CargaSelNumCriteriosTecnicos() As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim nCodigo As Integer
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  max(cCodCriterioTec) as codigo from LogSelCriteriosTecnicos"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If IsNull(rs!codigo) Then
       nCodigo = 1
       Else
       nCodigo = rs!codigo + 1
    End If
     CargaSelNumCriteriosTecnicos = nCodigo
End Function

Public Function ValidaSelNumCriteriosProceso(ByVal pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim nCodigo As Integer
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  count(cCodCriterioTec) as Cuenta from LogSelCriteriosProceso where nLogSelNro  =" & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
     ValidaSelNumCriteriosProceso = rs!Cuenta
End Function

'Devuelve un RecordSet con LogSelPublica
Public Function CargaSelPublica(ByVal pnSelNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then

        sSql = "SELECT sp.cPersCod, p.cPersNombre, Convert(varchar(12),sp.dLogSelPubIni,103), " & _
            "       Convert(varchar(12),sp.dLogSelPubFin,103) " & _
            " FROM LogSelPublica SP LEFT JOIN " & vsCentralPer & "Persona P ON sp.cPersCod = p.cPersCod " & _
            " WHERE sp.nLogSelNro = " & pnSelNro & " "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelPublica = rs
End Function


'Devuelve un RecordSet con LogSelCotiza
Public Function CargaSelCotiza(ByVal pnTpoCarga As TpoCargaSelCot, _
ByVal pnSelNro As Long, Optional ByVal pnSelCotNro As Long = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = SelCotNroCotPersona Then
            sSql = "SELECT sc.nLogSelCotNro, sc.cPersCod, p.cPersNombre, p.cPersDireccDomicilio " & _
                " FROM LogSelCotiza SC LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE sc.nLogSelNro = " & pnSelNro & " "
        ElseIf pnTpoCarga = SelCotPersona Then
            If pnSelCotNro = 0 Then
                sSql = "SELECT sc.cPersCod, p.cPersNombre, p.cPersDireccDomicilio " & _
                    " FROM LogSelCotiza SC LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
                    " WHERE sc.nLogSelNro = " & pnSelNro & " "
            Else
                sSql = "SELECT sc.cPersCod, p.cPersNombre, p.cPersDireccDomicilio " & _
                    " FROM LogSelCotiza SC LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
                    " WHERE sc.nLogSelNro = " & pnSelNro & " AND sc.nLogSelCotNro = '" & pnSelCotNro & "'"
            End If
        ElseIf pnTpoCarga = SelCotGarantia Then
            sSql = "SELECT IsNull(sc.nLogSelCotGarTpo,'') nLogSelCotGarTpo, IsNull(c.cConsDescripcion,'') cConsDescripcion, " & _
                "       IsNull(sc.cLogSelCotGarNro,'') cLogSelCotGarNro, IsNull(sc.cLogSelCotGarComentario,'') cLogSelCotGarComentario " & _
                " FROM LogSelCotiza SC LEFT JOIN Constante C ON sc.nLogSelCotGarTpo = c.nConsValor " & _
                "           And c.nConsCod = " & gPersGarantia & " " & _
                " WHERE sc.nLogSelNro = " & pnSelNro & " AND sc.nLogSelCotNro = '" & pnSelCotNro & "'"
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCotiza = rs
End Function

'Devuelve un RecordSet con LogSelCotDetalle
Public Function CargaSelCotDetalle(ByVal psTpoCarga As TpoCargaSelCotDet, _
ByVal pnSelNro As Integer, ByVal pnSelCotNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = SelCotDetUnRegistro Then
            sSql = "SELECT scd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       scd.nLogSelCotDetCantidad, scd.nLogSelCotDetPrecio, " & _
                "       (scd.nLogSelCotDetCantidad * scd.nLogSelCotDetPrecio) Total" & _
                " FROM LogSelCotDetalle SCD LEFT JOIN BienesServicios BS ON scd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE scd.nLogSelCotNro = '" & pnSelCotNro & "' AND scd.nLogSelNro = '" & pnSelNro & "'" & _
                " ORDER BY scd.cBSCod "
        ElseIf psTpoCarga = SelCotDetUnRegistroCotiza Then
            sSql = "SELECT scd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       sd.nLogSelDetCantidad, sd.nLogSelDetPrecio, " & _
                "       (sd.nLogSelDetCantidad * sd.nLogSelDetPrecio) Total, " & _
                "       scd.nLogSelCotDetCantidad, scd.nLogSelCotDetPrecio, " & _
                "       (scd.nLogSelCotDetCantidad * scd.nLogSelCotDetPrecio) TotalProv" & _
                " FROM LogSelCotDetalle SCD LEFT JOIN BienesServicios BS ON scd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                "       JOIN LogSelDetalle SD ON scd.nLogSelNro = sd.nLogSelNro " & _
                " WHERE scd.nLogSelCotNro = '" & pnSelCotNro & "' AND scd.nLogSelNro = '" & pnSelNro & "'" & _
                "       AND sd.cBSCod = scd.cBSCod "
        ElseIf psTpoCarga = SelCotDetUnRegEvalua Then
            sSql = "SELECT ad.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ad.nLogSelDetCantidad, ad.nLogSelDetPrecio, " & _
                "       (ad.nLogSelDetCantidad * ad.nLogSelDetPrecio) Total " & _
                " FROM LogSeleccion S JOIN LogSelDetalle AD ON s.nLogSelNro = ad.nLogSelNro " & _
                "       LEFT JOIN BienesServicios BS ON ad.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE s.nLogSelNro = " & pnSelNro & " " & _
                " ORDER BY ad.cBSCod "
        ElseIf psTpoCarga = SelCotDetUnRegEvaluaTodos Then
            sSql = "SELECT sc.nLogSelCotNro, scd.cBSCod, " & _
                "       scd.nLogSelCotDetCantidad, scd.nLogSelCotDetPrecio, " & _
                "       (scd.nLogSelCotDetCantidad * scd.nLogSelCotDetPrecio) Total" & _
                " FROM LogSelCotiza SC JOIN LogSelCotDetalle SCD ON sc.nLogSelNro = scd.nLogSelNro " & _
                "       AND sc.nLogSelCotNro = scd.nLogSelCotNro " & _
                " WHERE sc.nLogSelNro = " & pnSelNro & " " & _
                " ORDER BY sc.nLogSelNro, sc.nLogSelCotNro, scd.cBSCod  "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCotDetalle = rs
End Function

'Devuelve un RecordSet con datos de Locación de Servicios
Public Function CargaConLocacion(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    'NOTA.
    'CAMBIAR 2 x CONSTANTE DE CONTRATO DE LOCACION DE SERVICIOS
    If oConec.AbreConexion() Then
         sSql = " SELECT rhc.cPersCod, p.cPersNombre, " & _
            "       Convert(varchar(12), rhcd.dRHContratoInicio, 103), " & _
            "       Convert(varchar(12), rhcd.dRHContratoFin, 103), " & _
            "       IsNull(rhs.nRHSueldoMonto,0) nRHSueldoMonto " & _
            " FROM RHContrato RHC LEFT JOIN " & vsCentralPer & "Persona P ON rhc.cPersCod = p.cPersCod" & _
            "       LEFT JOIN RHContratoDet RHCD ON rhc.cPersCod = rhcd.cPersCod " & _
            "           AND rhc.cRHContratoNro = rhcd.cRHContratoNro " & _
            "       LEFT JOIN RHSueldo RHS ON rhc.cPersCod = rhs.cPersCod " & _
            "           AND rhs.dRHSueldoFecha Between RHCD.dRHContratoInicio " & _
            "           AND RHCD.dRHContratoFin " & _
            " WHERE rhc.cRHContratoNro = (select max(cRHContratoNro) " & _
            "                           from RHContrato where cPersCod = rhc.cPersCod) " & _
            "       AND rhcd.dRHContratoFecha = (select max(dRHContratoFecha) " & _
            "                           from RHContratoDet where cPersCod = rhc.cPersCod)" & _
            "       AND rhc.nRHContratoTpo = " & 1 & " " & _
            "       AND DateDiff(d, RHCD.dRHContratoFin, '" & Format(pdFecha, gsFormatoFecha) & "') < 0 " & _
            "       " & _
            " ORDER BY p.cPersNombre"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaConLocacion = rs
End Function

Public Function CargaLogSelTipoBs() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select Descripcion, cCodTipoBS    from LogSelTipoBs "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelTipoBs = rs
End Function

Public Function CargaLogSeleccionDetalle(psLogCodseleccion As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select Descripcion, cCodTipoBS    from LogSelTipoBs "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSeleccionDetalle = rs
End Function


Public Function LogSelTipoSeleccion(ByVal psTpoBS As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select t.Descripcion, c.cCodTipoSelec,c.nRangoIni,c.nRangoFin,c.cCodTipoBS     from LogSelConfProceso c,LogSelTipoSeleccion t where  " _
            & " cCodTipoBS = '" & psTpoBS & "' And c.cCodTipoSelec = t.cCodTipoSelec "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set LogSelTipoSeleccion = rs

End Function

Public Function LogSelConfProceso(ByVal psTpoBS As Integer, psTipoSeleccion As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  nRangoIni,nRangoFin from LogSelConfProceso where cCodTipoBS = " & psTpoBS & " and cCodTipoSelec = " & psTipoSeleccion & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set LogSelConfProceso = rs
End Function
Public Function LogSeleccionCod(ByVal psPeriodo As Integer) As Long
    Dim oConec As DConecta
    Dim sCodigo As String
    Dim nCodigo As Long
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  max(nLogSelNro) as nLogSelNro from logseleccionP where  nPeriodo = " & psPeriodo & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If IsNull(rs!nLogSelNro) = True Then
       sCodigo = Trim(Str(psPeriodo) + "00001")
    Else
       sCodigo = rs!nLogSelNro + 1
    End If
    nCodigo = Val(sCodigo)
    LogSeleccionCod = nCodigo
End Function

Public Function LogSeleccionLista(ByVal psPeriodo As Integer) As Recordset
    Dim oConec As DConecta
    Dim sCodigo As String
    Dim nCodigo As Long
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  CAST(s.nLogSelNro as varchar(10)) as nLogSelNro ,s.nPeriodo,t.Descripcion,p.Descripcion, " _
               & " nLogSelEstado = CASE    WHEN nLogSelEstado = 0 THEN  'INICIADO' " _
               & " WHEN nLogSelEstado = 1 THEN 'EVALUACION TECNICA'  WHEN nLogSelEstado = 2 THEN 'EVALUACION ECONOMICA' " _
               & " WHEN nLogSelEstado = 3 THEN 'FIN DE EVALUACION'   WHEN nLogSelEstado = 4 THEN 'CERRADO' " _
               & " WHEN nLogSelEstado = 5 THEN 'CANCELADO' END  " _
               & " from LogseleccionP s,LogSelTipoBs t,LogSelTipoSeleccion p " _
               & " where s.cCodTipoBS = t.cCodTipoBS and " _
               & " s.cCodTipoSelec = p.cCodTipoSelec and  s.nPeriodo = " & psPeriodo & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set LogSeleccionLista = rs
End Function

Public Function LogSeleccionListaConPara(ByVal psPeriodo As Integer, ByVal pnEstado As Integer, pnMoneda As Integer) As Recordset
    Dim oConec As DConecta
    Dim sCodigo As String
    Dim nCodigo As Long
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  CAST(s.nLogSelNro as varchar(10)) as nLogSelNro ,s.nPeriodo,t.Descripcion,p.Descripcion, " _
               & " nLogSelEstado = CASE    WHEN nLogSelEstado = 0 THEN  'INICIADO' " _
               & " WHEN nLogSelEstado = 1 THEN 'EVALUACION TECNICA'  WHEN nLogSelEstado = 2 THEN 'EVALUACION ECONOMICA' " _
               & " WHEN nLogSelEstado = 3 THEN 'FIN DE EVALUACION'   WHEN nLogSelEstado = 4 THEN 'CERRADO' " _
               & " WHEN nLogSelEstado = 5 THEN 'CANCELADO' END  " _
               & " from LogseleccionP s,LogSelTipoBs t,LogSelTipoSeleccion p " _
               & " where s.cCodTipoBS = t.cCodTipoBS and " _
               & " s.cCodTipoSelec = p.cCodTipoSelec and  s.nPeriodo = " & psPeriodo & " and  " _
               & " nLogSelEstado  = '" & pnEstado & "' and nLogSelMoneda = '" & pnMoneda & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set LogSeleccionListaConPara = rs
End Function




Public Function LogSeleccionListaProveedores(ByVal pnLogSelNro As Long) As Recordset
    Dim oConec As DConecta
    Dim sCodigo As String
    Dim nCodigo As Long
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  SP.cPersCod,cPersNombre ,1 as Nivel from LogSelProveedor SP,Persona P" _
               & " Where SP.cPersCod = P.cPersCod  and nLogSelNro =" & pnLogSelNro & " " _
               & " Union " _
               & " Select  '0' as cPersCod ,'LISTADO DE PROVEEDORES' as  cPersNombre ,0 as Nivel  order by  Nivel asc "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set LogSeleccionListaProveedores = rs
End Function


Public Function InsertaSeleccionLog(pnSelNro As Long, psPeriodo As Integer, psTipoBien As Integer, psTipoSeleccion As Integer, psMoneda As Integer, psRangIni As Long, psRangoFin As Long, psFechaConvo As String, psActualiza As String, psDescripcion As String, psNumeroCotizacion As String, pnClaseProceso As Integer) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionLog = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = " Insert Into logseleccionP(nLogSelNro,nPeriodo,cCodTipoBS,cCodTipoSelec, nLogSelEstado," _
             & " nLogSelMoneda, cUltimaActualizacion,nRangoIni,NRangoFin,cFechaConvoSeleccion,cDescripcionProceso,nLogSelNumeroCot,nLogSelTipoProceso) Values(" _
             & " " & pnSelNro & "," & psPeriodo & " ," & psTipoBien & "," & psTipoSeleccion & ",0, " _
             & " " & psMoneda & ",'" & sActualiza & "'," & psRangIni & "," & psRangoFin & ",'" & psFechaConvo & "','" & psDescripcion & "','" & psNumeroCotizacion & "'," & pnClaseProceso & ")"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionLog = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetalle Method")
End Function
Public Function ActualizaSeleccionLog(pnSelNro As Long, pnPeriodo As Integer, pnTipoBien As Integer, pnTipoSeleccion As Integer, pnMoneda As Integer, pnRangIni As Long, pnRangoFin As Long, psFechaConvo As String, psActualiza As String, psDescripcion, psNumeroCotizacion As String, pnClaseProceso As Integer) As Integer
On Error GoTo ActualizaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActualizaSeleccionLog = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "Update LogseleccionP set  " _
               & " nperiodo =" & pnPeriodo & ", " _
               & " cCodTipoBS  = " & pnTipoBien & ", " _
               & " cCodTipoSelec = " & pnTipoSeleccion & ", " _
               & " nLogSelEstado = 0 ," _
               & " nLogSelMoneda  = " & pnMoneda & ", " _
               & " cFechaConvoSeleccion = '" & psFechaConvo & "'," _
               & " nRangoIni = " & pnRangIni & "," _
               & " nRangoFin  = " & pnRangoFin & "," _
               & " cUltimaActualizacion = '" & psActualiza & "', " _
               & " cDescripcionProceso ='" & psDescripcion & "', " _
               & " nLogSelNumeroCot ='" & psNumeroCotizacion & "'," _
               & " nLogSelTipoProceso = " & pnClaseProceso & " " _
               & " where  nLogSelNro = " & pnSelNro & " "
        oConec.Ejecutar sSql
        oConec.CierraConexion
        Set oConec = Nothing
        ActualizaSeleccionLog = 0
        Exit Function
ActualizaSeleccionLogErr:
        Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetalle Method")
End Function
Public Function InsertaSeleccionConfPuntajes(pnnLogSelNro As Long, pnCodTipoEvaluacion As Integer, pnPuntajeMinimo As Integer, pnPuntajeMaximo As Integer, pnPesoPonderado As Double, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionConfPuntajes = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelConfigPuntajes " _
                & " (nLogSelNro,cCodTipoEvaluacion,nPuntajeMinimo,nPuntajeMaximo,nPesoPonderado,cUltimaActualizacion) " _
                & " values (" & pnnLogSelNro & "," & pnCodTipoEvaluacion & "," & pnPuntajeMinimo & "," & pnPuntajeMaximo & "," & pnPesoPonderado & ",'" & psActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionConfPuntajes = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetalle Method")
End Function
Public Function EliminaSeleccionConfPuntajes(pnLogSelNro As Long) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionConfPuntajes = 1
    oConec.AbreConexion
    sSql = "Delete from  LogSelConfigPuntajes where  nLogSelNro = " & pnLogSelNro & "  "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionConfPuntajes = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdquisi:EliminaSeleccionConfPuntajes Method")
End Function






Public Function ActuzalizaSeleccionConfPuntajes(pnnLogSelNro As Long, pnCodTipoEvaluacion As Integer, pnPuntajeMinimo As Integer, pnPuntajeMaximo As Integer, pnPesoPonderado As Double, psActualiza As String) As Integer
On Error GoTo ActuzalizaSeleccionConfPuntajesErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActuzalizaSeleccionConfPuntajes = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "Update  LogSelConfigPuntajes  set " _
        & " cCodTipoEvaluacion = " & pnCodTipoEvaluacion & " , " _
        & " nPuntajeMinimo =" & pnPuntajeMinimo & " , " _
        & " nPuntajeMaximo = " & pnPuntajeMaximo & " , " _
        & " nPesoPonderado = " & pnPesoPonderado & ", " _
        & " cUltimaActualizacion ='" & psActualiza & "' " _
        & " where nLogSelNro = " & pnnLogSelNro & "   "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActuzalizaSeleccionConfPuntajes = 0
    Exit Function
ActuzalizaSeleccionConfPuntajesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActuzalizaSeleccionConfPuntajes Method")
End Function
Public Function ActualizaSelEvaluacionEco(pnLogSelNro As Long, psCodPers As String, psCodBien As String, pnPuntaje As Double, psFechaEvaluacionEco As String) As Integer
On Error GoTo ActualizaSelevaluacionEcoErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim nPuntajeTecnico As Double
    Set rs = New ADODB.Recordset
    Set oConec = New DConecta
    ActualizaSelEvaluacionEco = 1
    oConec.AbreConexion
    nPuntajeTecnico = 0
    psFechaEvaluacionEco = Format(oConec.GetFechaHoraServer, "yyyymmdd") + Format(oConec.GetHoraServer, "hhmmss")
    sSql = "Select  sum(nPuntaje) as PuntajeTecnico from  LogSelCalificacionTecnica where  nLogSelNro = " & pnLogSelNro & " and  cPersCod = '" & psCodPers & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    If IsNull(rs!PuntajeTecnico) = True Then
            nPuntajeTecnico = 0
    Else
            nPuntajeTecnico = rs!PuntajeTecnico
    End If
          sSql = "Update  LogSelCotDetalle  Set " _
          & " nPuntajeEconomico =" & pnPuntaje & ", " _
          & " nPuntajeTecnico =" & nPuntajeTecnico & ", " _
          & " cfechaEvaluacionEco = '" & psFechaEvaluacionEco & "' " _
          & " where nLogSelNro = " & pnLogSelNro & " and  cBSCod ='" & psCodBien & "'  and  cPersCod ='" & psCodPers & "'  and indicadorParticipa =1 "
          oConec.Ejecutar sSql
          sSql = "Update  LogSelCotDetalle  Set " _
          & " nPuntajeTecnico = 0," _
          & " nPuntajeEconomico  = 0, " _
          & " cfechaEvaluacionEco  = '" & psFechaEvaluacionEco & "' " _
          & " where nLogSelNro = " & pnLogSelNro & " and  cBSCod ='" & psCodBien & "'  and  cPersCod ='" & psCodPers & "'  and indicadorParticipa = 0 "
          oConec.Ejecutar sSql
          oConec.CierraConexion
    Set oConec = Nothing
    ActualizaSelEvaluacionEco = 0
    Exit Function
ActualizaSelevaluacionEcoErr:
    Call RaiseError(MyUnhandledError, "DLogAdquisi:ActualizaSelevaluacionEco Method")
End Function

Public Function InsertaSeleccionComite(pnLogSelNro As Long, psCodPer As String, pncargo As Integer, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionComite = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelComite " _
                & " (nLogSelNro,cPersCod,nCargoCod,cUltimaActualizacion) " _
                & " values (" & pnLogSelNro & ",'" & psCodPer & "'," & pncargo & ",'" & psActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionComite = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionComite Method")
End Function
Public Function InsertaSeleccionConfigBienes(pnLogSelNro As Long, pscBSCod As String, psDescripcion As String, pnCantidad As Integer, pnUnidad As Integer, pnPrecio As Double, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionConfigBienes = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelDetalle " _
               & "(nLogSelNro,cBSCod,nLogSelDetCantidad,cDescripcionBien," _
               & " nLogSelUnidad,nLogSelDetPrecio,cUltimaActualizacion) " _
               & " Values( " _
               & " " & pnLogSelNro & ",'" & pscBSCod & "'," & pnCantidad & " ,'" & psDescripcion & "'," _
               & " " & pnUnidad & "," & pnPrecio & ",'" & psActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionConfigBienes = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionConfigBienes Method")
End Function

Public Function InsertaLogSeleccionCotDetalle(pnLogSelNro As Long, pscBSCod As String, psCodPer As String, _
            pnCantidad As Integer, pnUnidad As Integer, pnPrecio As Double, _
            pnIndicadorParticipa As Integer, psActualiza As String) As Integer
On Error GoTo InsertaLogSeleccionCotDetalleErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaLogSeleccionCotDetalle = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelCotDetalle " _
               & "(nLogSelNro,cBSCod,cPersCod,nLogSelCotDetCantidad," _
               & " nLogSelUnidad,nLogSelCotDetPrecio,IndicadorParticipa,cUltimaActualizacion) " _
               & " Values( " _
               & " " & pnLogSelNro & ",'" & pscBSCod & "','" & psCodPer & "' ," & pnCantidad & "," _
               & " " & pnUnidad & "," & pnPrecio & "," & pnIndicadorParticipa & ",'" & psActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaLogSeleccionCotDetalle = 0
    Exit Function
InsertaLogSeleccionCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaLogSeleccionCotDetalle Method")
End Function


Public Function EliminaLogSeleccionCotDetalle(pnLogSelNro As Long, psCodPersona As String) As Integer
On Error GoTo EliminaLogSeleccionCotDetalleErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaLogSeleccionCotDetalle = 1
    oConec.AbreConexion
        sSql = "Delete from LogSelCotDetalle where  nLogSelNro =" & pnLogSelNro & "  and cPersCod ='" & psCodPersona & "'  "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaLogSeleccionCotDetalle = 0
    Exit Function
EliminaLogSeleccionCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaLogSeleccionCotDetalle Method")
End Function

Public Function InsertaSeleccionCriteriosProceso(pnLogSelNro As Long, pnCodCriterioTec As Integer, pnPuntajeCritProceso As Integer, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionCriteriosProceso = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelCriteriosProceso " _
               & "(nLogSelNro,cCodCriterioTec,nPuntajeCritProceso,cUltimaActualizacion)" _
               & " Values( " _
               & " " & pnLogSelNro & "," & pnCodCriterioTec & "," & pnPuntajeCritProceso & ",'" & sActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionCriteriosProceso = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionCriteriosProceso Method")
End Function

Public Function InsertaSeleccionCalificacionTecnica(pnLogSelNro As Long, pnCodCriterioTec As Integer, psCodProveedor As String, _
                pnPuntaje As Double, psDescripcion As String, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionCalificacionTecnica = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "Insert into  LogSelCalificacionTecnica ( " _
               & " nLogSelNro,cCodCriterioTec,cPersCod, " _
               & " nPuntaje,cDescripcionEvaluacion,cUltimaActualizacion) values (" _
               & " " & pnLogSelNro & "," & pnCodCriterioTec & ",'" & psCodProveedor & "', " _
               & " " & pnPuntaje & ",'" & psDescripcion & "','" & psActualiza & "' )"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionCalificacionTecnica = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionCalificacionTecnica Method")
End Function

Public Function InsertaSeleccionProveedores(pnLogSelNro As Long, psCodPer As String, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionProveedores = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into LogSelProveedor " _
               & "(nLogSelNro,cPersCod,cUltimaActualizacion)" _
               & " Values( " _
               & " " & pnLogSelNro & ",'" & psCodPer & "','" & sActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionProveedores = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionProveedores Method")
End Function


Public Function EliminaSeleccionCriteriosProceso(pnLogSelNro As Long) As Integer
On Error GoTo EliminaSeleccionCriteriosProcesoErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionCriteriosProceso = 1
    oConec.AbreConexion
    sSql = "Delete from LogSelCriteriosProceso where  nLogSelNro =" & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionCriteriosProceso = 0
    Exit Function
EliminaSeleccionCriteriosProcesoErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionCriteriosProceso Method")
End Function
Public Function EliminaSeleccionCalificacionTecnica(pnLogSelNro As Long, psCodProveedor As String) As Integer
On Error GoTo EliminaSeleccionCalificacionTecnicaErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionCalificacionTecnica = 1
    oConec.AbreConexion
    sSql = "Delete from LogSelCalificacionTecnica where  nLogSelNro = " & pnLogSelNro & " and  cPersCod ='" & psCodProveedor & "'"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionCalificacionTecnica = 0
    Exit Function
EliminaSeleccionCalificacionTecnicaErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionCalificacionTecnica Method")
End Function



Public Function EliminaSeleccionProveedores(pnLogSelNro As Long) As Integer
On Error GoTo EliminaSeleccionProveedoresErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionProveedores = 1
    oConec.AbreConexion
    sSql = "Delete from  LogSelProveedor where  nLogSelNro =" & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionProveedores = 0
    Exit Function
EliminaSeleccionProveedoresErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionProveedores Method")
End Function




Public Function InsertaLogSelReferencia(pnLogSelNro As Long, pnLogSelPeriodo As Integer, pnLogSelTpoReq As Integer, _
                pnConsolidado As Integer, psDescripcionConsol As String, _
                pnmesIni As String, pnmesFin As String, pscategoriabien, psActualiza As String) As Integer
On Error GoTo InsertaLogSelReferenciaErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaLogSelReferencia = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "Insert Into LogSelReferencia (" _
              & " nLogSelNro,nLogSelPeriodo, " _
              & " nLogSelTpoReq,nConsolidado,SDescripcionConsol, " _
              & " nMesIni,nMesFin,sCategoriaBien,cUltimaActualizacion ) " _
              & " values(" & pnLogSelNro & "," & pnLogSelPeriodo & " ," _
              & " " & pnLogSelTpoReq & "," & pnConsolidado & ",'" & psDescripcionConsol & "'," _
              & " '" & pnmesIni & "','" & pnmesFin & "','" & pscategoriabien & "','" & psActualiza & "') "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaLogSelReferencia = 0
    Exit Function
InsertaLogSelReferenciaErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaLogSelReferencia Methodo")
End Function
Public Function EliminaLogSelReferencia(pnLogSelNro As Long) As Integer
On Error GoTo EliminaLogSelReferenciaErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaLogSelReferencia = 1
    oConec.AbreConexion
    sSql = " Delete from  LogSelReferencia where nLogSelNro  = " & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaLogSelReferencia = 0
    Exit Function
EliminaLogSelReferenciaErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaLogSelReferencia Methodo")
End Function

Public Function EliminaLogSelMantCriterios(nCodCriterioTec As Integer) As Integer
On Error GoTo EliminaLogSelMantCriteriosErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaLogSelMantCriterios = 1
    oConec.AbreConexion
    sSql = "Delete from LogSelCriteriosTecnicos where cCodCriterioTec =" & nCodCriterioTec & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaLogSelMantCriterios = 0
    Exit Function
EliminaLogSelMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaLogSelMantCriterios Methodo")
End Function
Public Function InsertaLogSelMantCriterios(pnCodCriterioTec As Integer, psDescripcion As String, pnPuntajeDefault As Double, psUltimaActualizacion As String) As Integer
On Error GoTo InsertaLogSelMantCriteriosErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaLogSelMantCriterios = 1
    oConec.AbreConexion
    sActualiza = Left(psUltimaActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psUltimaActualizacion, 15)
    sSql = " Insert Into LogSelCriteriosTecnicos " _
           & " (cCodCriterioTec,cDescripcion,nPuntajeDefault,cUltimaActualizacion ) " _
           & "  values(" & pnCodCriterioTec & ",'" & psDescripcion & "'," & pnPuntajeDefault & ",'" & sActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaLogSelMantCriterios = 0
    Exit Function
InsertaLogSelMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaLogSelMantCriterios Methodo")
End Function
Public Function ActualizaEstadoProveedor(pnLogSeleccion As Long, psProveedor As String, pnEstado As Integer) As Integer
On Error GoTo ActualizaEstadoProveedorErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActualizaEstadoProveedor = 1
    oConec.AbreConexion
    sSql = " Update LogSelProveedor  set  nLogSelProEstado = " & pnEstado & "  " _
    & " where nLogSelNro  = " & pnLogSeleccion & " and  cPersCod ='" & psProveedor & "'"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActualizaEstadoProveedor = 0
    Exit Function
ActualizaEstadoProveedorErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:ActualizaEstadoProveedor Methodo")
End Function
Public Function ActualizaEstadoProcesoSeleccion(pnLogSeleccion As Long, pnEstado As Integer) As Integer
On Error GoTo ActualizaEstadoProcesoSeleccionErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActualizaEstadoProcesoSeleccion = 1
    oConec.AbreConexion
    sSql = " Update LogSeleccionP  set  nLogSelEstado = " & pnEstado & "  " _
    & " where nLogSelNro  = " & pnLogSeleccion & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActualizaEstadoProcesoSeleccion = 0
    Exit Function
ActualizaEstadoProcesoSeleccionErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:ActualizaEstadoProcesoSeleccion Methodo")
End Function
Public Function InsertaMotAnulacion(pnLogSeleccion As Long, pnMotAnulacion As Integer, psActualizacion) As Integer
On Error GoTo InsertaMotAnulacionErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaMotAnulacion = 1
    oConec.AbreConexion
    sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "insert Into LogSelCancelacion (nLogSelNro,nLogSelCodMotCancel,cUltimaActualizacion) " _
    & " Values (" & pnLogSeleccion & "," & pnMotAnulacion & ",'" & psActualizacion & "') "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaMotAnulacion = 0
    Exit Function
InsertaMotAnulacionErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaMotAnulacion Methodo")
End Function

Public Function ActualizaLogSelMantCriterios(pnCodCriterioTec As Integer, psDescripcion As String, pnPuntajeDefault As Double, psUltimaActualizacion As String) As Integer
On Error GoTo ActualizaLogSelMantCriteriosErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActualizaLogSelMantCriterios = 1
    oConec.AbreConexion
    sActualiza = Left(psUltimaActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psUltimaActualizacion, 15)
    sSql = " Update LogSelCriteriosTecnicos set " _
         & " cDescripcion   = '" & psDescripcion & "', " _
         & " nPuntajeDefault = " & pnPuntajeDefault & "," _
         & " cUltimaActualizacion ='" & sActualiza & "' " _
         & " where cCodCriterioTec = " & pnCodCriterioTec & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActualizaLogSelMantCriterios = 0
    Exit Function
ActualizaLogSelMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:ActualizaLogSelMantCriterios Methodo")
End Function



Public Function CargaSelReferencia(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select nLogSelNro ,nLogSelPeriodo,nLogSelTpoReq," _
               & " nConsolidado ,SDescripcionConsol,nMesIni, " _
               & " nMesFin , sCategoriaBien, cUltimaActualizacion " _
               & " from  LogSelReferencia where nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelReferencia = rs
End Function

Public Function CargaLogSelCriteriosProceso(pnLogSelNro As Long, Tipo As Integer, Optional ByVal psCodProvedor As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        Select Case Tipo
        Case 1
              sSql = "Select  p.cCodCriterioTec, t.cDescripcion,nPuntajeCritProceso,p.cUltimaActualizacion from  LogSelCriteriosProceso p,LogSelCriteriosTecnicos t " _
              & " where  p.cCodCriterioTec = t.cCodCriterioTec and nLogSelNro = " & pnLogSelNro & " "
        Case 2
              sSql = "Select  p.cCodCriterioTec, t.cDescripcion,nPuntajeDefault from  LogSelCriteriosProceso p,LogSelCriteriosTecnicos t " _
              & " where  p.cCodCriterioTec = t.cCodCriterioTec and nLogSelNro = " & pnLogSelNro & " "
        Case 3
             
              sSql = "Select c.cCodCriterioTec ,p.cDescripcion ,nPuntajeCritProceso,c.nPuntaje,cDescripcionEvaluacion,c.cUltimaActualizacion " _
                     & " From LogSelCalificacionTecnica c,LogSelCriteriosTecnicos p  ,LogSelCriteriosProceso d  " _
                     & " Where c.cCodCriterioTec = P.cCodCriterioTec  and " _
                     & " c.nLogSelNro = d.nLogSelNro and c.cCodCriterioTec = d.cCodCriterioTec and " _
                     & " cPersCod = '" & psCodProvedor & "' and c.nLogSelNro =" & pnLogSelNro & " "
                     
                     
        End Select
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelCriteriosProceso = rs
End Function
Public Function CargaLogSelProveedores(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "Select  SP.cPersCod,cPersNombre,SP.cUltimaActualizacion from LogSelProveedor SP,Persona P" _
                   & "  where SP.cPersCod = p.cPersCod and  nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelProveedores = rs
End Function

Public Function CargaLogSelDescripcionProceso(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "Select  rtrim(bs.Descripcion) +' -  '+rtrim(ts.Descripcion) as ctipo,cDescripcionProceso,nLogSelNumeroCot,nLogSelEstado,  " _
                & " nLogSelTipoProceso, nLogSelDescProceso =  CASE   WHEN nLogSelTipoProceso = 1 THEN 'NORMAL'    WHEN nLogSelTipoProceso = 2 THEN 'DIRECTO'  END " _
                & " from logseleccionP s,LogSelTipoBs bs,LogSelTipoSeleccion ts " _
                & " where  s.cCodTipoBS = bs.cCodTipoBS and  s.cCodTipoSelec = ts.cCodTipoSelec and nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelDescripcionProceso = rs
End Function
Public Function CargaLogSelEstadoProceso(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = " Select  nLogSelEstado from logseleccionP " _
               & " Where   nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaLogSelEstadoProceso = rs!nLogSelEstado
End Function
Public Function CargaLogSelComiteRep(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select  L.nCargoCod,cPersNombre,cDescripcion,cPersIDnro from " _
               & " LogSelComite C, Persona P,LogSelcargos L ,PERSID D " _
               & " where c.cPersCod = p.cPersCod  and  L.nCargoCod = c.nCargoCod and " _
               & " D.cPersCod = C.cPersCod AND " _
               & " nLogSelNro = " & pnLogSelNro & " order by c.nCargoCod asc  "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelComiteRep = rs
End Function
Public Function CargaLogSelProvBienes(pnLogSelNro As Long, psCodProveedor As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select  ld.cBSCod,cBSDescripcion ,nLogSelCotDetCantidad,nLogSelCotDetPrecio,nLogSelCotDetPrecio*nLogSelCotDetCantidad as SubTotal," _
               & " nPuntajeEconomico * (select  nPesoPonderado from logselconfigpuntajes where nLogSelNro = " & pnLogSelNro & " and cCodTipoEvaluacion = 1) as Tecnico, " _
               & " nPuntajeTecnico * (select  nPesoPonderado from logselconfigpuntajes where nLogSelNro = " & pnLogSelNro & " and cCodTipoEvaluacion = 2) as Economico, " _
               & " nPuntajeEconomico * (select  nPesoPonderado from logselconfigpuntajes where nLogSelNro = " & pnLogSelNro & " and cCodTipoEvaluacion = 1) + " _
               & " nPuntajeTecnico * (select  nPesoPonderado from logselconfigpuntajes where nLogSelNro = " & pnLogSelNro & " and cCodTipoEvaluacion = 2) as PuntajeTotal " _
               & " from logselcotdetalle ld, Bienesservicios bs where bs.cBSCod = ld.cBSCod  and nLogSelNro = " & pnLogSelNro & " AND " _
               & " cPersCod = '" & psCodProveedor & "' AND nIndicaGanador = 'SI' "
 
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelProvBienes = rs
End Function





Public Function ValidaLogSelCalificacionTecnica(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  count(*) as Cuenta from LogSelCalificacionTecnica where nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    ValidaLogSelCalificacionTecnica = rs!Cuenta
End Function
Public Function ValidaLogSelCalificacionEconomica(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  count(*) as Cuenta from LogSelCotDetalle where nLogSelNro = " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    ValidaLogSelCalificacionEconomica = rs!Cuenta
End Function
Public Function CargaLogSelResumenEvaluacionTecnica(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "execute spLogSelEvaluaciontecnica  " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelResumenEvaluacionTecnica = rs
End Function
Public Function CargaLogSelResumenCuadro(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "execute spLogSelResumenTotal  " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelResumenCuadro = rs
End Function

Public Function ActualizaLogSelganador(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    ActualizaLogSelganador = 0
    If oConec.AbreConexion() Then
        sSql = "execute spLogSelIndiganador  " & pnLogSelNro & " "
        oConec.Ejecutar (sSql)
        oConec.CierraConexion
    End If
    ActualizaLogSelganador = 1
End Function



Public Function CargaLogSelResumenEvaluacionEconomica(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "execute spLogSelEvaluacionEconomica  " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelResumenEvaluacionEconomica = rs
End Function

Public Function CargaLogSelActaBuenaPro(pnLogSelNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "execute SPLogActaBuenaPro  " & pnLogSelNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelActaBuenaPro = rs
End Function



Public Function CargaLogSelNombresProveedor(psPersCod As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  cPersNombre from Persona where  cPersCod = '" & psPersCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
        If rs.EOF = True Then
        CargaLogSelNombresProveedor = "No Encontrado"
        Else
        CargaLogSelNombresProveedor = rs!cPersNombre
        End If
    
     
End Function



Public Function EliminaSeleccionComite(pnLogSelNro As Long) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionComite = 1
    oConec.AbreConexion
     sSql = "delete from  LogSelComite " _
                & " where  nLogSelNro =" & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionComite = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionComite Method")
End Function

Public Function EliminaSeleccionConfigBienes(pnLogSelNro As Long) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionConfigBienes = 1
    oConec.AbreConexion
    sSql = "Delete from logSelDetalle where  nLogSelNro = " & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionConfigBienes = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionConfigBienes Method")
End Function


Public Function TipoUsado(pnCodCriterioTec As Integer) As Boolean
    On Error GoTo TipoUsadoErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlC = " Select cCodCriterioTec From LogSelCriteriosProceso where cCodCriterioTec = " & pnCodCriterioTec & ""
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.EOF And rs.BOF Then
            TipoUsado = False
        Else
            TipoUsado = True
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
TipoUsadoErr:
    Call RaiseError(MyUnhandledError, "DLogAdquisi:TipoUsado Method")
End Function

Public Function CargaEstadoProveedor(pnLogSelNro As Long, psProveedor As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select Estado = case nLogSelProEstado when 0 then  'CALIFICADO' when 1 then 'DESCALIFICADO' else 'Sin Estado' End from LogSelProveedor where nLogSelNro =" & pnLogSelNro & " and  cPersCod  ='" & psProveedor & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaEstadoProveedor = rs!Estado
End Function

Public Function ValidaEvalTecProvedor(pnLogSelNro As Long, psProveedor As String) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select count(*) as Total from LogSelCalificacionTecnica where nLogSelNro =" & pnLogSelNro & " and  cPersCod  ='" & psProveedor & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    ValidaEvalTecProvedor = rs!Total
End Function

Public Function CargaMotRechazos() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "select   Descripcion,nLogSelCodMotCancel  from LogSelMotCancel where nStatusRegistro = 0 "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaMotRechazos = rs
End Function
Public Function CargaLogSelMotCancel(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select  nLogSelCodMotCancel from LogSelCancelacion where  nLogSelNro  = " & pnLogSelNro & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaLogSelMotCancel = IIf(IsNull(rs!nLogSelCodMotCancel) = True, 0, rs!nLogSelCodMotCancel)
End Function

Public Function CargaLogSelEstadoDesc(nestadoProc As Integer) As String
         If nestadoProc = SelEstProcesoIniciado Then
               CargaLogSelEstadoDesc = "INICIADO"
         ElseIf nestadoProc = SelEstProcesoEvaluacionTec Then
               CargaLogSelEstadoDesc = "EVALUACION TECNICA"
         ElseIf nestadoProc = SelEstProcesoEvaluacionEco Then
               CargaLogSelEstadoDesc = "EVALUACION ECONOMICA"
         ElseIf nestadoProc = SelEstProcesoFinEvaluacion Then
               CargaLogSelEstadoDesc = "FIN DE EVALUACION"
         ElseIf nestadoProc = SelEstProcesoCerrado Then
               CargaLogSelEstadoDesc = "CERRADO"
         ElseIf nestadoProc = SelEstProcesoCancelado Then
               CargaLogSelEstadoDesc = "ANULADO"
         End If
End Function
Public Function CargaSelTipoComite() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  cDescripcioTipoComite,cCodTipoComite   from logSelTipoComite"
       
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelTipoComite = rs
End Function

Public Function EliminaSeleccionTipoComite(pnTipoComite As Long) As Integer
On Error GoTo EliminaSeleccionProveedoresErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaSeleccionTipoComite = 1
    oConec.AbreConexion
    sSql = "delete  from logselcomitepredeter  where cCodTipoComite= " & pnTipoComite & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaSeleccionTipoComite = 0
    Exit Function
EliminaSeleccionProveedoresErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaSeleccionTipoComite Method")
End Function

Public Function InsertaSeleccionTipoComite(pnTipoComite As Long, psCodPer As String, pnCodcargo As Integer, psActualiza As String) As Integer
On Error GoTo InsertaSeleccionLogErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaSeleccionTipoComite = 1
    oConec.AbreConexion
    'cperscod      cCodTipoComite cUltimaActualizacion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
        sSql = "insert into logselcomitepredeter " _
               & "(cCodTipoComite,cPersCod,ncargoCod,cUltimaActualizacion)" _
               & " Values( " _
               & " " & pnTipoComite & ",'" & psCodPer & "'," & pnCodcargo & ",'" & sActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaSeleccionTipoComite = 0
    Exit Function
InsertaSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaSeleccionTipoComite Method")
End Function

Public Function CargaLogSelComitePre(pnTipComite As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
           sSql = "Select s.cPersCod,p.cPersNombre,c.cDescripcion + replicate(' ',100) +  convert(char(1),s.nCargoCod),s.cUltimaActualizacion  " _
           & " from logselcomitepredeter s,persona p, logselcargos c " _
           & " where  s.cPersCod = p.cPersCod and c.nCargoCod = s.nCargoCod And cCodTipoComite = " & pnTipComite & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaLogSelComitePre = rs
End Function

Public Function CargalogselcriteriosTecnicos(pnCodCCriterio As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select  nPuntajeDefault from logselcriteriostecnicos  where  nStatusRegistro =1 and cCodCriterioTec = " & pnCodCCriterio & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargalogselcriteriosTecnicos = rs
End Function

Public Function ValidaLogSelEmpate(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  cBSCod,nIndicaGanador,count(*) as Cant from logselcotdetalle " _
               & " where  nLogSelNro =" & pnLogSelNro & " and  nIndicaGanador ='SI'" _
               & " group by cBSCod,nIndicaGanador having count(*) > 1  "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If rs.EOF = True Then
        ValidaLogSelEmpate = 0
    Else
        ValidaLogSelEmpate = rs!Cant
    End If
    
    
End Function

Public Function CargalogSelEmpates(pnLogSelNro As Long, pnEmpate As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "spLogSelEmpate " & pnLogSelNro & "," & pnEmpate & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargalogSelEmpates = rs
End Function

Public Function InsertaLogDesenmpate(pnLogSelNro As Long, pscBSCod As String, psCodPer As String, psDescripcion As String, psActualiza As String) As Integer
On Error GoTo InsertaLogSeleccionCotDetalleErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    InsertaLogDesenmpate = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
    
    sSql = "insert into logseldesempate(nLogSelNro,cPersCod,cBSCod,Comentarios,cUltimaActualizacion)" _
                & " values (" & pnLogSelNro & ",'" & psCodPer & "','" & pscBSCod & "','" & psDescripcion & "','" & psActualiza & "')"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    InsertaLogDesenmpate = 0
    Exit Function
InsertaLogSeleccionCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:InsertaLogDesenmpate Method")
End Function

Public Function EliminaLogDesenmpate(pnLogSelNro As Long) As Integer
On Error GoTo InsertaLogSeleccionCotDetalleErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    EliminaLogDesenmpate = 1
    oConec.AbreConexion
    sSql = "Delete from logseldesempate where  nLogSelNro = " & pnLogSelNro & " "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminaLogDesenmpate = 0
    Exit Function
InsertaLogSeleccionCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:EliminaLogDesenmpate Method")
End Function

Public Function ActualizaLogDesenmpate(pnLogSelNro As Long, pscBSCod As String, psCodPer As String, psIndicaGanador As String, psActualiza As String) As Integer
On Error GoTo ActualizaLogDesenmpateErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    ActualizaLogDesenmpate = 1
    oConec.AbreConexion
    sActualiza = Left(psActualiza, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualiza, 15)
    sSql = "Update logselcotdetalle set  " _
            & " nIndicaGanador = '" & psIndicaGanador & "' , " _
            & " cUltimaActualizacion =  '" & psActualiza & "'" _
            & " where  nLogSelNro =" & pnLogSelNro & "    and  " _
            & " cBSCod  =   '" & pscBSCod & "'and cPersCod = '" & psCodPer & "'   "
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActualizaLogDesenmpate = 0
    Exit Function
ActualizaLogDesenmpateErr:
    Call RaiseError(MyUnhandledError, "DLogAdqui:ActualizaLogDesenmpate Method")
End Function

Public Function ValidaLogSelDesempate(pnLogSelNro As Long) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  count(*) as cant from logseldesempate where nLogSelNro = " & pnLogSelNro & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If rs.EOF = True Then
        ValidaLogSelDesempate = 0
    Else
        ValidaLogSelDesempate = rs!Cant
    End If
End Function
Public Function CargalogSelPlantilla(pnTipoPlantilla As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select cDescripcionPlantilla,nCodPlantilla  from LogSelPlantilla where  nTipoPlanilla  =" & pnTipoPlantilla & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargalogSelPlantilla = rs
End Function

Public Function CargalogSelDescPlantilla(pnTipoPlantilla As Long, pnCodPlanilla) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select cNombrePlantilla from LogSelPlantilla where  nTipoPlanilla  = " & pnTipoPlantilla & " and nCodPlantilla =  " & pnCodPlanilla & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If rs.EOF = True Then
         CargalogSelDescPlantilla = ""
    Else
         CargalogSelDescPlantilla = IIf(IsNull(rs!cNombrePlantilla) = True, "", rs!cNombrePlantilla)
    End If
End Function
Public Function CargaDirRUCPersona(cCodPersona As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT  cPersDireccDomicilio ,cPersIDnro FROM PERSONA A , PERSID B " _
                & " WHERE A.cPersCod = B.cPersCod and  cPersIDTpo  = 2 and a.cperscod ='" & cCodPersona & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaDirRUCPersona = rs
End Function





