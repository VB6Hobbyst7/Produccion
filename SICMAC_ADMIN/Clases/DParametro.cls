VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaDatos(Optional ByVal psProducto As String = "", Optional ByVal pnOrden As Integer = 1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
' pnOrden ==> 0 = (cParamDesc) // 1 = (nParamVar)
On Error GoTo ErrorRecuperaDatos
    sSql = "Select nParamVar,cParamDesc,nParamValor From ColocParametro " & _
           "Where nParamVar like '" & psProducto & "%' " & _
           "Order by " & IIf(pnOrden = 1, "cParamDesc", "nParamVar")
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaDatos:
    Err.Raise Err.Number, "Recupera Datos", Err.Description

End Function

Public Function RecuperaValorParametro(ByVal pnCodigo As Long) As Double
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaValorParametro
    sSql = "Select nParamValor from ColocParametro Where nParamVar = " & pnCodigo
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    'RecuperaValorParametro = R!nParamValor
    'WIOR 20130923 ********************************
    If R.RecordCount > 0 Then
        RecuperaValorParametro = R!nParamValor
    Else
        RecuperaValorParametro = 0
    End If
    'WIOR 20130923 ********************************
    Exit Function

ErrorRecuperaValorParametro:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    

End Function

Public Function ExisteParametro(ByVal psVariable As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select * from ColocParametro where nParamVar = '" & psVariable & "'"
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        ExisteParametro = True
    Else
        ExisteParametro = False
    End If
    
End Function

Public Sub NuevoParametro(ByVal psVariable As String, ByVal psDesc As String, ByVal pnValor As Double)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorNuevoParametro
        
    sSql = "INSERT INTO ColocParametro(nParamVar,cParamDesc,nParamValor)"
    sSql = sSql & " VALUES('" & psVariable & "','" & psDesc & "'," & Format(pnValor, "#0.00") & ")"
    
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoParametro:
    Err.Raise Err.Number, "Nuevo Parametro", Err.Description
End Sub

Public Sub ModificarParametro(ByVal psVariable As String, ByVal psDesc As String, ByVal pnValor As Double)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificarParametro
    
    sSql = "UPDATE ColocParametro SET cParamDesc='" & Replace(psDesc, "'", "''") & "', nParamValor = " & Format(pnValor, "#0.00")
    sSql = sSql & " Where nParamVar = " & psVariable
    
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificarParametro:
    Err.Raise Err.Number, "Modifcar Parametro", Err.Description
End Sub

Public Sub EliminarParametro(ByVal psVariable As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorEliminarParametro
        
        sSql = "DELETE ColocParametro "
        sSql = sSql & " Where nParamVar = " & psVariable
        
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarParametro:
    Err.Raise Err.Number, "Eliminar Parametro", Err.Description
End Sub


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
