VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Function ValidaDatos(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String) As String
Dim nLongTemp As Integer
Dim oConecta As DConecta
Dim sSQL As String
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
            nLongTemp = 11
    End Select
    
    'Valida Existencia de Descendientes
    Set oConecta = New DConecta
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrorExisteLinea
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select cLineaCred from ColocLineacredito Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSQL)
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
    Optional ByVal pbPreferencial As Boolean = False) As String
Dim oLineaCredito As DLineaCredito
Dim i, j, k As Integer
Dim Abv1, Abv2 As String

    On Error GoTo ErrorNuevaLinea
    NuevaLinea = ""
    Set oLineaCredito = New DLineaCredito
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
    
    Set oLineaCredito = New DLineaCredito
    
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
                    For j = 1 To Len(psDescProduc)
                        If Mid(psDescProduc, j, 1) = " " Then
                            k = j + 1
                            Exit For
                        End If
                    Next j
                    If k <> -1 Then
                        Abv2 = Mid(psDescProduc, k, 3)
                    End If
                    If Not ExisteLinea(Mid(psLineaCred, 1, 5)) Then
                        Call oLineaCredito.NuevaLineaCredito(Mid(psLineaCred, 1, 5), Abv1 & "-" & Abv2, 1, 0, 0, 0#, 0#, psPersCod)
                    End If
                    
        End Select
    Next i
    
    Call oLineaCredito.NuevaLineaCredito(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, , pbPreferencial)
    
    Set oLineaCredito = Nothing
    Exit Function

ErrorNuevaLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ModificarLinea(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, _
ByVal psPersCod As String, Optional ByVal pbPreferencia As Boolean = False) As String
Dim oLineaCredito As DLineaCredito

    On Error GoTo ErrorModificar
    ModificarLinea = ""
    ModificarLinea = ValidaDatos(pnLongitud, psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod)
    If ModificarLinea <> "" Then
        Exit Function
    End If
    
    Set oLineaCredito = New DLineaCredito
    If pnLongitud <> 1 Then
        Call oLineaCredito.ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, oLineaCredito.RecuperaInstitucion(psLineaCred), pbPreferencia)
    Else
        Call oLineaCredito.ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, pbPreferencia)
    End If
    Set oLineaCredito = Nothing

    Exit Function

ErrorModificar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Private Function ExisteReferencia(ByVal psLineaCred As String) As String
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta
    
    On Error GoTo ErrorExisteReferencia
    ExisteReferencia = ""
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select cLineaCred from ColocLineacreditoTasa Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta
    
    On Error GoTo ErrorExisteReferenciaConCredito
    ExisteReferenciaConCredito = ""
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select cLineaCred from Colocaciones Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta
    
    On Error GoTo ErrorExisteSaldoLinea
    ExisteSaldoLinea = ""
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select cLineaCred from ColocLineaCreditoSaldo Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim oLineaCredito As DLineaCredito
    
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
    Set oLineaCredito = New DLineaCredito
    Call oLineaCredito.EliminaLineaCredito(psLineaCred)
    Set oLineaCredito = Nothing

    Exit Function

ErrorEliminaLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Private Function ExisteTasa(ByVal psLineaCred As String, ByVal psColocLinCredTasaTpo As String) As Boolean
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta
    
    On Error GoTo ErrorExisteTasa
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select cLineaCred from ColocLineacreditoTasa Where cLineaCred = '" & psLineaCred
    sSQL = sSQL & "' And nColocLinCredTasaTpo = " & psColocLinCredTasaTpo
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim oLinea As DLineaCredito

    On Error GoTo ErrorNuevaTasa
    'Nivel de Creacion de Tasas
    If Len(cLineaCred) <> 11 Then
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
    
    Set oLinea = New DLineaCredito
    Call oLinea.NuevaLineaCreditoTasas(cLineaCred, cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    Set oLinea = Nothing
    
    Exit Function
    
ErrorNuevaTasa:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function ModificarTasa(ByVal cLineaCred As String, ByVal cColocLinCredTasaTpo As String, ByVal nTasaIni As Double, ByVal nTasaFin As Double) As String
Dim oLinea As DLineaCredito
    On Error GoTo ErrorModificarTasa
    
    ModificarTasa = ValidaDatosTasa(cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    If ModificarTasa <> "" Then
        Exit Function
    End If
    
    Set oLinea = New DLineaCredito
    Call oLinea.ActualizarLineaTasas(cLineaCred, cColocLinCredTasaTpo, nTasaIni, nTasaFin)
    Set oLinea = Nothing

    Exit Function

ErrorModificarTasa:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function EliminarTasa(ByVal cLineaCred As String, ByVal cColocLinCredTasaTpo As String) As String
Dim oLinea As DLineaCredito

    On Error GoTo ErrorEliminarTasa
    EliminarTasa = ""
    Set oLinea = New DLineaCredito
    Call oLinea.EliminaTasasLinea(cLineaCred, cColocLinCredTasaTpo)
    Exit Function

ErrorEliminarTasa:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
