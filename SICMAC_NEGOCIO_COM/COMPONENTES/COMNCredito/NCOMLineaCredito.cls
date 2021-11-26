VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Function ValidaDatos(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String) As String
Dim nLongTemp As Integer
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer

    ValidaDatos = ""
    Select Case pnLongitud
        Case 1
            nLongTemp = 2
        Case 2
            nLongTemp = 5
        Case 3
            nLongTemp = 6
        Case 4
            nLongTemp = 9
        Case 5
            nLongTemp = 13 '11(Nueva Longitud 13)
    End Select
    
    'Valida Existencia de Descendientes
    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    For I = 2 To pnLongitud
'        Select Case I
'            Case 2
'                sSql = "Select cLineaCred from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 2) & "'"
'                ValidaDatos = "No se ha definido Fondo para esta Linea"
'            Case 3
'                sSql = "Select cLineaCred from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 5) & "'"
'                ValidaDatos = "No se ha definido el Sub Fondo y la Moneda "
'            Case 4
'                sSql = "Select cLineaCred from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 6) & "'"
'                ValidaDatos = "No se ha definido el Plazo para esta Linea"
'            Case 5
'                sSql = "Select cLineaCred from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 9) & "'"
'                ValidaDatos = "No se ha definido el Producto para esta Linea"
'        End Select
'        Set R = oConecta.CargaRecordSet(sSql)
'        If R.RecordCount = 0 Then
'            R.Close
'            oConecta.CierraConexion
'            Set oConecta = Nothing
'            Exit Function
'        Else
'            ValidaDatos = ""
'        End If
'    Next I
'    Set R = Nothing
'    oConecta.CierraConexion
    Set oConecta = Nothing
    
    'Valida Longitud
    If Len(Trim(psLineaCred)) <> nLongTemp Then
        ValidaDatos = "La Longitud del Codigo de la Linea de Credito No es Correcta"
        Exit Function
    End If
        
    'Valida Plazos
    If pnPlazoMax <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Plazo Maximo debe ser mayor que Cero"
        Exit Function
    End If
    If pnPlazoMin <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Plazo Minimo debe ser mayor que Cero"
        Exit Function
    End If
    
    If pnPlazoMin > pnPlazoMax Then
        ValidaDatos = "El Plazo Minimo no debe ser mayor el Plazo Maximo"
        Exit Function
    End If
    
    'Valida Monto Maximo y Minimo
    If pnMontoMin <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Monto Minimo debe ser mayor que Cero"
        Exit Function
    End If
    If pnMontoMax <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Monto Maximo debe ser mayor que Cero"
        Exit Function
    End If
    
    If pnMontoMin > pnMontoMax Then
        ValidaDatos = "El Monto Minimo no debe ser mayor el Monto Maximo"
        Exit Function
    End If
    
    'Valida Institucion
    If Len(Trim(psPersCod)) = 0 And pnLongitud = 1 Then
        ValidaDatos = "Debe Ingresar La Institucion"
        Exit Function
    End If
    
    
End Function

Public Function ExisteLinea(ByVal psLineaCred As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorExisteLinea
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from ColocLineacredito Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteLinea = True
    Else
        ExisteLinea = False
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorExisteLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function NuevaLinea(ByVal pnLongitud As Integer, ByVal psLineaCred As String, _
    ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, _
    ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, _
    ByVal psPersCod As String, ByVal psDescFondo As String, ByVal psDescSubFondo As String, _
    ByVal psDescProduc As String, Optional ByVal psAbrevFondo As String = "", Optional ByVal psAbrevSubFondo As String = "", _
    Optional ByVal pbPreferencial As Boolean = False, _
    Optional ByVal pMatAgencias As Variant) As String

Dim oLineaCredito As COMDCredito.DCOMLineaCredito
Dim i, J, k As Integer
Dim Abv1, Abv2 As String

    On Error GoTo ErrorNuevaLinea
    NuevaLinea = ""
    Set oLineaCredito = New COMDCredito.DCOMLineaCredito
    psLineaCred = psLineaCred & oLineaCredito.Correlativo(psLineaCred)
    Set oLineaCredito = Nothing
    NuevaLinea = ValidaDatos(pnLongitud, psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod)
    If NuevaLinea <> "" Then
        Exit Function
    End If
    
    If ExisteLinea(psLineaCred) Then
        NuevaLinea = "Linea de Credito ya Existe"
        Exit Function
    End If
    
    Set oLineaCredito = New COMDCredito.DCOMLineaCredito
    
    For i = 1 To 4
        Select Case i
            Case 1 'Crear el Fondo
                If Not ExisteLinea(Mid(psLineaCred, 1, 2)) Then
                    Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 2), psDescFondo, 1, 0, 0, 0#, 0#, psPersCod, psAbrevFondo)
                End If
            Case 2 'Crear el Sub fondo
                If Not ExisteLinea(Mid(psLineaCred, 1, 5)) Then
                    Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 5), psDescSubFondo, 1, 0, 0, 0#, 0#, psPersCod, psAbrevSubFondo)
                End If
            Case 3 'Crear en el Plazo Escojido
                If Not ExisteLinea(Mid(psLineaCred, 1, 6)) Then
                    If Mid(psLineaCred, 6, 1) = "1" Then
                        Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 6), psDescSubFondo & " - CORTO PL", 1, 0, 0, 0#, 0#, psPersCod)
                    Else
                        Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 6), psDescSubFondo & " - LARGO PL", 1, 0, 0, 0#, 0#, psPersCod)
                    End If
                End If
            Case 4 'Crear el Producto
                    k = -1
                    Abv1 = Mid(psDescProduc, 1, 3)
                    For J = 1 To Len(psDescProduc)
                        If Mid(psDescProduc, J, 1) = " " Then
                            k = J + 1
                            Exit For
                        End If
                    Next J
                    If k <> -1 Then
                        Abv2 = Mid(psDescProduc, k, 3)
                    End If
                    If Not ExisteLinea(Mid(psLineaCred, 1, 5)) Then
                        Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 5), Abv1 & "-" & Abv2, 1, 0, 0, 0#, 0#, psPersCod)
                    End If
                    
        End Select
    Next i
    
    Call oLineaCredito.NuevaLineaCredito(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, , pbPreferencial, pMatAgencias)
    
    Set oLineaCredito = Nothing
    Exit Function

ErrorNuevaLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ModificarLinea(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, _
                                ByVal psPersCod As String, Optional ByVal pbPreferencia As Boolean = False, Optional ByVal pMatAgencias As Variant) As String

Dim oLineaCredito As COMDCredito.DCOMLineaCredito

    On Error GoTo ErrorModificar
    ModificarLinea = ""
    ModificarLinea = ValidaDatos(pnLongitud, psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod)
    If ModificarLinea <> "" Then
        Exit Function
    End If
    
    Set oLineaCredito = New COMDCredito.DCOMLineaCredito
    If pnLongitud <> 1 Then
        Call oLineaCredito.ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, oLineaCredito.RecuperaInstitucion(psLineaCred), pbPreferencia, pMatAgencias)
    Else
        Call oLineaCredito.ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, pbPreferencia, pMatAgencias)
    End If
    Set oLineaCredito = Nothing

    Exit Function

ErrorModificar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Private Function ExisteReferencia(ByVal psLineaCred As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorExisteReferencia
    ExisteReferencia = ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from ColocLineacreditoTasa Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteReferencia = "No se puede Ejecutar el Proceso porque posee Tasas Definidas"
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteReferencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Private Function ExisteReferenciaConCredito(ByVal psLineaCred As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorExisteReferenciaConCredito
    ExisteReferenciaConCredito = ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from Colocaciones Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteReferenciaConCredito = "La Linea esta Asignada a un Credito"
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteReferenciaConCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Private Function ExisteSaldoLinea(ByVal psLineaCred As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorExisteSaldoLinea
    ExisteSaldoLinea = ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from ColocLineaCreditoSaldo Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteSaldoLinea = "La Linea Posee Saldo y no puede ser Eliminada"
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteSaldoLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function EliminaLinea(ByVal psLineaCred As String) As String
Dim oLineaCredito As COMDCredito.DCOMLineaCredito
    
    On Error GoTo ErrorEliminaLinea
    EliminaLinea = ExisteReferencia(psLineaCred)
    If EliminaLinea <> "" Then
        Exit Function
    End If
    EliminaLinea = ExisteReferenciaConCredito(psLineaCred)
    If EliminaLinea <> "" Then
        Exit Function
    End If
    EliminaLinea = ExisteSaldoLinea(psLineaCred)
    If EliminaLinea <> "" Then
        Exit Function
    End If
    Set oLineaCredito = New COMDCredito.DCOMLineaCredito
    Call oLineaCredito.EliminaLineaCredito(psLineaCred)
    Set oLineaCredito = Nothing

    Exit Function

ErrorEliminaLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Private Function ExisteTasa(ByVal psLineaCred As String, ByVal psColocLinCredTasaTpo As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorExisteTasa
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from ColocLineacreditoTasa Where cLineaCred = '" & psLineaCred
    sSql = sSql & "' And nColocLinCredTasaTpo = " & psColocLinCredTasaTpo
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteTasa = True
    Else
        ExisteTasa = False
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteTasa:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Private Function ValidaDatosTasa(ByVal cColocLinCredTasaTpo As String, ByVal nTasaIni As Double, ByVal nTasaFin As Double) As String
    
    ValidaDatosTasa = ""
    
    'Valida Tipo de Tasa
    If Len(Trim(cColocLinCredTasaTpo)) = 0 Then
        ValidaDatosTasa = "Ingrese Tipo de Tasa"
        Exit Function
    End If
    
    'Valida Montos
    If nTasaIni <= 0 Then
        ValidaDatosTasa = "La Tasa Inicial debe ser Mayor que Cero"
        Exit Function
    End If
    If nTasaFin <= 0 Then
        ValidaDatosTasa = "La Tasa Final debe ser Mayor que Cero"
        Exit Function
    End If
    
    'Valida Montos
    If nTasaIni > nTasaFin Then
        ValidaDatosTasa = "Tasa Inicial no debe ser Mayor a la Tasa Final"
        Exit Function
    End If
End Function
Public Function NuevaTasa(ByVal cLineaCred As String, ByVal cColocLinCredTasaTpo As String, ByVal nTasaIni As Double, ByVal nTasaFin As Double) As String
Dim oLinea As COMDCredito.DCOMLineaCredito

    On Error GoTo ErrorNuevaTasa
    'Nivel de Creacion de Tasas
    If Len(cLineaCred) <> 13 Then   '11(Nueva Longitud)
        NuevaTasa = "No puede Crear Tasas a este Nivel"
        Exit Function
    End If
    NuevaTasa = ValidaDatosTasa(cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    If NuevaTasa <> "" Then
        Exit Function
    End If
    If ExisteTasa(cLineaCred, cColocLinCredTasaTpo) Then
        NuevaTasa = "Tasa ya Existe"
        Exit Function
    End If
    
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Call oLinea.NuevaLineaCreditoTasas(cLineaCred, cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    Set oLinea = Nothing
    
    Exit Function
    
ErrorNuevaTasa:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function ModificarTasa(ByVal cLineaCred As String, ByVal cColocLinCredTasaTpo As String, ByVal nTasaIni As Double, ByVal nTasaFin As Double) As String
Dim oLinea As COMDCredito.DCOMLineaCredito
    On Error GoTo ErrorModificarTasa
    
    ModificarTasa = ValidaDatosTasa(cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    If ModificarTasa <> "" Then
        Exit Function
    End If
    
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Call oLinea.ActualizarLineaTasas(cLineaCred, cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    Set oLinea = Nothing

    Exit Function

ErrorModificarTasa:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function EliminarTasa(ByVal cLineaCred As String, ByVal cColocLinCredTasaTpo As String) As String
Dim oLinea As COMDCredito.DCOMLineaCredito

    On Error GoTo ErrorEliminarTasa
    EliminarTasa = ""
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Call oLinea.EliminaTasasLinea(cLineaCred, cColocLinCredTasaTpo)
    Exit Function

ErrorEliminarTasa:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'Public Sub CambioLineaCredito(ByVal psCtaCod, _
'            ByVal pcLineaCredOrigen As String, _
'            ByVal pnSaldoCapOrigen As Double, _
'            ByVal pcLineaCredDestino As String, _
'            ByVal pnSaldoCapDestino As Double, _
'            ByVal pdHoy As Date, _
'            ByVal psCodAge As String, _
'            ByVal psCodUser As String, _
'            Optional psMovnroTemp As String = "", _
'            Optional ByVal pbInicioTrans As Boolean = False, Optional ByRef psMensajeValidacion As String = "")
'Dim oBase As COMDCredito.DCOMCredActBD
'Dim pConnBase As COMDCredito.DCOMCredActBD
'Dim nEstadoCred As Integer
'Dim nEstadoCredTemp As Integer
'Dim R As ADODB.Recordset
'Dim oCred As COMDCredito.DCOMCredito
'Dim oCredN As COMNCredito.NCOMCredito
'Dim oCalend As COMDCredito.DCOMCalendario
'Dim nTransacc As Long
''Dim sLineaCred As String
''Dim nMontoColocado As Double
''Dim dFecPend As Date
'Dim nDiasAtraso As Integer
'Dim i, k As Integer
'Dim nNroCalen As Integer
'Dim nMontoGasto As Double
'Dim sMovNro As String
'Dim nmovnro As Long
'Dim nConsCred As String
'Dim nConsCred2 As String
'Dim pnPlazo As Integer
'Dim bTran As Boolean
''Dim oFunciones As New COMNContabilidad.NCOMContFunciones
'Dim dFechaTran As Date
''Dim nIntPend As Double
''Dim nMontoPago As Double
''Dim nMontoPago_2 As Double
'Dim nNroCalPar As Integer
'Dim nMivivienda As Integer
'Dim MatGastosCred As Variant
'Dim NumregGastosCred As Integer
'Dim MatGastosCuota As Variant
'Dim NumRegGastosCuota As Integer
'Dim nBuenPagador As Integer
'Dim nPrestamo As Double
''Dim CapitalPagado As Double
'Dim nDiasAtrasoMov As Integer
''Dim opeITFChequeEfect As String
'Dim MatCalend As Variant
'Dim MatCalendDistrib As Variant
'
'Dim psSaldoKFecha As String, psIntCompFecha As String, psIntMorFecha As String, psGastoFecha As String
'Dim psPenalidadFecha As String, psTotalFecha As String
'Dim psCapitalVencido, psIntCompVencido, psIntCompVencVencido, psIntMoratorioVencido, psTotalFechaVencido As String
'
'Dim oITF As COMDConstSistema.FCOMITF
'Dim oGastos As COMNCredito.NCOMGasto
'
'Dim MatGastosCancelacion As Variant
'Dim nNumGastosCancel As Integer
'
'Dim pnMonto As Double
'Dim psMetLiquid As String
'Dim pnPrepago As Integer
'
'Dim oExec As COMDExecSP.DCOMExecSP
'
'    On Error GoTo ErrorCambioLineaCredito
'
'    Set pConnBase = Nothing
'    bTran = False
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaProducto(psCtaCod)
'    Set oCred = Nothing
'    nEstadoCred = R!nPrdEstado
'    nEstadoCredTemp = R!nPrdEstado
'    nTransacc = R!nTransacc     'nro de transacción
'    R.Close
'
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaColocaciones(psCtaCod)
'    Set oCred = Nothing
'    nPrestamo = R!nMontoCol
'    R.Close
'    Set R = Nothing
'
'    'Definir Codigo de Operacion de Cambio de Linea
'    Select Case nEstadoCred
'        'Si es Credito refinanciado
'        Case gColocEstRefMor
'            nConsCred = gCredCambioLineaRefNor
'            nConsCred2 = gCredCambioLineaRefVen
'        Case gColocEstRefNorm
'            nConsCred = gCredCambioLineaRefNor
'        Case gColocEstRefVenc
'            nConsCred = gCredCambioLineaRefVen
'        'si es Credito Normal
'        Case gColocEstVigMor
'            nConsCred = gCredCambioLineaVigNor
'            nConsCred2 = gCredCambioLineaVigVen
'        Case gColocEstVigNorm
'            nConsCred = gCredCambioLineaVigNor
'        Case gColocEstVigVenc
'            nConsCred = gCredCambioLineaVigVen
'        'si es credito Judicial
'        Case gColocEstRecVigJud
'            nConsCred = gCredCambioLineaRecJud
'        Case gColocEstRecVigCast
'            nConsCred = gCredCambioLineaRecCas
'    End Select
'
'    If nConsCred = "" Then
'        'MsgBox "No se ha definido la operación correctamente " & vbCrLf & _
'               "Consulte con la Oficina de T.I", vbInformation, "AVISO"
'        psMensajeValidacion = "No se ha definido la operación correctamente. Consulte con la Oficina de T.I"
'        Exit Sub
'    End If
'
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
'    Set oCred = Nothing
'    pnPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
'    R.Close
'    Set R = Nothing
'
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaColocacCred(psCtaCod)
'    Set oCred = Nothing
'    nNroCalen = R!nNroCalen
'    nNroCalPar = R!nNroCalPar
'    nMivivienda = IIf(IsNull(R!bMiVivienda), 0, R!bMiVivienda)
'    nBuenPagador = IIf(IsNull(R!nCalPago), 0, R!nCalPago)
'    nDiasAtrasoMov = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
'    psMetLiquid = IIf(IsNull(R!cMetLiquidacion), "", R!cMetLiquidacion)
'    pnPrepago = 0
'    R.Close
'    Set R = Nothing
'
'    '********************************************************************************
'    Set oCredN = New COMNCredito.NCOMCredito
'    MatCalend = oCredN.RecuperaMatrizCalendarioPendiente(psCtaCod)
'    MatCalendDistrib = oCredN.CrearMatrizparaAmortizacion(MatCalend)
'
'    'Calculo a la fecha de CAPITAL, INTERES, MORA Y EL TOTAL DE LA DEUDA
'    If UBound(MatCalend) > 0 Then
'        psSaldoKFecha = Format(oCredN.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
'        If nEstadoCred = 2022 Or nEstadoCred = 2032 Then
'            psCapitalVencido = Format(oCredN.MatrizCapitalVencido(MatCalend, pdHoy), "#0.00")
'            psIntCompVencido = Format(oCredN.MatrizIntCompVencido(MatCalend, pdHoy), "#0.00")
'            psIntCompVencVencido = Format(oCredN.MatrizIntCompVencVencido(MatCalend, pdHoy), "#0.00")
'            psIntMoratorioVencido = Format(oCredN.ObtenerMoraVencida(pdHoy, MatCalend), "#0.00")
'            psTotalFechaVencido = Format(CDbl(psCapitalVencido) + CDbl(psIntCompVencido) + CDbl(psIntCompVencVencido) + CDbl(psIntMoratorioVencido), "#0.00")
'        Else
'            psCapitalVencido = Format(0, "#0.00")
'            psIntCompVencido = Format(0, "#0.00")
'            psIntCompVencVencido = Format(0, "#0.00")
'            psIntMoratorioVencido = Format(0, "#0.00")
'            psTotalFechaVencido = Format(0, "#0.00")
'        End If
'        psIntCompFecha = Format(oCredN.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy), "#0.00")
'        psIntMorFecha = Format(oCredN.ObtenerMoraVencida(pdHoy, MatCalend), "#0.00")
'        psGastoFecha = Format(0, "#0.00")
'        psPenalidadFecha = "0.00"
'        psTotalFecha = Format(CDbl(psSaldoKFecha) + CDbl(psIntCompFecha) + CDbl(psGastoFecha) + CDbl(psIntMorFecha) + CDbl(psPenalidadFecha), "#0.00")
'        Set oCredN = Nothing
'    Else
'        Set oCredN = Nothing
'    End If
'    '********************************************************************************
'
'    bTran = True
'    If pConnBase Is Nothing Then
'        Set oBase = New COMDCredito.DCOMCredActBD
'        If pbInicioTrans = False Then
'            Call oBase.dBeginTrans
'        End If
'    Else
'        Set oBase = pConnBase
'    End If
'
'    If psMovnroTemp <> "" Then
'        sMovNro = psMovnroTemp
'    Else
'        sMovNro = oBase.GeneraMovNro(pdHoy, psCodAge, psCodUser)
'    End If
'
'    'Insert Movimientos
'    If psMovnroTemp <> "" Then
'        nmovnro = oBase.dGetnMovNro(psMovnroTemp)
'        If nmovnro = 0 Then
'            Call oBase.dInsertMov(psMovnroTemp, nConsCred, "Cambio de Linea", gMovEstContabMovContable, gMovFlagVigente, False)
'            nmovnro = oBase.dGetnMovNro(psMovnroTemp)
'        End If
'    Else
'        Call oBase.dInsertMov(sMovNro, nConsCred, "Cambio de Linea", gMovEstContabMovContable, gMovFlagVigente, False)
'        nmovnro = oBase.dGetnMovNro(sMovNro)
'    End If
'
'    pnMonto = psTotalFecha
'    Set oCredN = New COMNCredito.NCOMCredito
'
'    'Registrando el total de la deuda en un movimiento
'    If nEstadoCred = 2022 Or nEstadoCred = 2032 Then
'        Call oBase.dInsertMovCol(nmovnro, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, oCredN.MatrizSaldoCapital(MatCalend, MatCalendDistrib) - psCapitalVencido, nEstadoCredTemp, False, , pnPrepago)
'        Call oBase.dInsertMovCol(nmovnro, nConsCred2, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, psCapitalVencido, nEstadoCredTemp, False, , pnPrepago)
'    Else
'        Call oBase.dInsertMovCol(nmovnro, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, oCredN.MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, , pnPrepago)
'    End If
'
'    'Amortizando Capital
'    If CDbl(psSaldoKFecha) > 0 Then
'        If CDbl(psCapitalVencido) > 0 Then
'        'If nEstadoCred = 2022 Or nEstadoCred = 2032 Then
'            'Inserta Detalle Movimiento Capital
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, 1, CDbl(psSaldoKFecha) - CDbl(psCapitalVencido), False)
'            'Inserta Detalle Movimiento Capital Vencido
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred2, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, 1, CDbl(psCapitalVencido), False)
'        Else
'            'Inserta Detalle Movimiento Capital
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, 1, CDbl(psSaldoKFecha), False)
'        End If
'    End If
'    'Amortizando Interes Compensatorio
'    If CDbl(psIntCompFecha) + CDbl(psIntMorFecha) > 0 Then
'        If CDbl(psIntCompVencido) + CDbl(psIntCompVencVencido) + CDbl(psIntMoratorioVencido) > 0 Then
'        'If nEstadoCred = 2022 Or nEstadoCred = 2032 Then
'            'Inserta Detalle Movimiento Interes Compensatorio
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompensatorio, 1, (CDbl(psIntCompFecha) - CDbl(psIntCompVencido) - CDbl(psIntCompVencVencido)) + (CDbl(psIntMorFecha) - CDbl(psIntMoratorioVencido)), False)
'            'Inserta Detalle Movimiento Interes Compensatorio - Vencido
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred2, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompensatorio, 1, CDbl(psIntCompVencido) + CDbl(psIntCompVencVencido) + CDbl(psIntMoratorioVencido), False)
'        Else
'            'Inserta Detalle Movimiento Interes Compensatorio
'            Call oBase.dInsertMovColDet(nmovnro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompensatorio, 1, CDbl(psIntCompFecha) + CDbl(psIntMorFecha), False)
'        End If
'    End If
'
''    'Amortizando Interes Moratorio
''    If CDbl(psIntMorFecha) > 0 Then
''        'Inserta Detalle Movimiento Interes Gracia
''        Call oBase.dInsertMovColDet(nmovnro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresMoratorio, 1, CDbl(psIntMorFecha), False)
''    End If
'
'    If pConnBase Is Nothing Then
'        If pbInicioTrans = False Then
'            Call oBase.dCommitTrans
'        End If
'    End If
'
'    If pConnBase Is Nothing Then
'        Set oBase = Nothing
'    End If
'
'    'Actualizamos el campo de linea de Credito
''    Dim oExec As COMDExecSP.DCOMExecSP
'    Set oExec = New COMDExecSP.DCOMExecSP
'    Call oExec.ExecuteSqlQuery("Update Colocaciones Set cLineaCred = '" & pcLineaCredDestino & "' Where cCtaCod = '" & psCtaCod & "'")
'    Set oExec = Nothing
'
'    'Actualizamos el Saldo de la Tabla CredSaldosAdeudo
'    Set oExec = New COMDExecSP.DCOMExecSP
'    Call oExec.ExecuteSqlQuery("Update ColocLineaCreditoSaldo Set nSaldoCap = " & pnSaldoCapOrigen & " Where cLineaCred = '" & pcLineaCredOrigen & "'")
'    Call oExec.ExecuteSqlQuery("Update ColocLineaCreditoSaldo Set nSaldoCap = " & pnSaldoCapDestino & " Where cLineaCred = '" & pcLineaCredDestino & "'")
'    Set oExec = Nothing
'
'    'Insertamos un registro en la tabla de Historial de la Linea de Credito
'    Dim oExecSP As COMDExecSP.DCOMExecSP
'    Set oExecSP = New COMDExecSP.DCOMExecSP
'    Call oExecSP.ExecuteStoreProcedure("Coloc_insert_ColocCredHistoricoCL_sp", psCtaCod, pcLineaCredOrigen, pnSaldoCapOrigen, pcLineaCredDestino, pnSaldoCapDestino, nmovnro)
'    Set oExecSP = Nothing
'
'    Exit Sub
'
'ErrorCambioLineaCredito:
'    If bTran Then
'        Call oBase.dRollbackTrans
'        Set oBase = Nothing
'    End If
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Sub




