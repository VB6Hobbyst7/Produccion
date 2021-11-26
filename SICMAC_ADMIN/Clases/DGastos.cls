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

Public Function RecuperaGastosCabecera() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaGastosCabecera
    sSql = "Select nColocConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor, nValor ,nMoneda from ColocConcepto"
    sSql = sSql & " Where convert(varchar(15),nColocConceptoCod) like '12%'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosCabecera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosCabecera:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function

Public Function RecuperaGastosAplicablesDesembolso() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaGastosAplicablesDesembolso
    sSql = "select * from ColocConcepto where nInicial >0 AND nAplicado in (" & gColocConceptoAplDesembolso & "," & gColocConceptoAplTodosD & "," & gColocConceptoAplTodosDC & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesDesembolso:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function

Public Function RecuperaGastosAplicablesCuotas() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaGastosAplicablesCuotas
    sSql = "select * from ColocConcepto where nFinal >0 AND nAplicado in (" & gColocConceptoAplCuota & "," & gColocConceptoAplTodosC & "," & gColocConceptoAplTodosDC & ") AND nDeCierre = 0 and nColocCred = 1"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosAplicablesCuotas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaGastosAplicablesCuotas:
    Err.Raise Err.Number, "Recupera Gastos", Err.Description
    
End Function


Public Sub NuevoGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorNuevoGasto
    sSql = "INSERT INTO ColocConcepto(nColocConceptoCod, cDescripcion, nAplicado, nInicial, nFinal, nTpoValor, nValor, nMoneda)"
    sSql = sSql & " VALUES(" & psColocConceptoCod & ",'" & psDescripcion & "'," & psAplicado & "," & Format(pnInicial, "#0.00") & "," & Format(pnFinal, "#0.00") & "," & psTpoValor & "," & Format(pnValor, "#0.00") & "," & psMoneda & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoGasto:
    Err.Raise Err.Number, "Nuevo Gasto", Err.Description
End Sub

Public Sub ActualizarGasto(ByVal psColocConceptoCod As String, ByVal psDescripcion As String, ByVal psAplicado As String, ByVal pnInicial As Double, ByVal pnFinal As Double, psTpoValor As String, ByVal pnValor As Double, ByVal psMoneda As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorActualizarGasto
    sSql = " UPDATE ColocConcepto SET "
    sSql = sSql & "cDescripcion = '" & psDescripcion & "',"
    sSql = sSql & "nAplicado = " & psAplicado & ","
    sSql = sSql & "nInicial = " & Format(pnInicial, "#0.00") & ","
    sSql = sSql & "nFinal = " & Format(pnFinal, "#0.00") & ","
    sSql = sSql & "nTpoValor = " & psTpoValor & ","
    sSql = sSql & "nValor = " & Format(pnValor, "#0.00") & ","
    sSql = sSql & "nMoneda = " & psMoneda & " "
    sSql = sSql & " WHERE nColocConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizarGasto:
    Err.Raise Err.Number, "Actualizar Gasto", Err.Description
End Sub

Public Function GastoAsignadoACredito(ByVal psColocConceptoCod As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    
    sSql = "Select * From ColocCalendDet Where nColocConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
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
    
    sSql = "Select * From MovColDet Where nColocConceptoCod = " & psColocConceptoCod
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarGasto
    sSql = " DELETE ColocConcepto "
    sSql = sSql & " WHERE nColocConceptoCod = " & psColocConceptoCod
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    On Error GoTo ErrorRecuperaGastoValorPenalidadCancelacion
    If pnMoneda = gMonedaNacional Then
        sSql = "Select nValor from ColocConcepto where nColocConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoSoles
    Else
        sSql = "Select nValor from ColocConcepto where nColocConceptoCod = " & gColocConceptoCodGastoPenalidadPagoAdelantadoDolares
    End If
    Set oConecta = New DConecta
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

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
