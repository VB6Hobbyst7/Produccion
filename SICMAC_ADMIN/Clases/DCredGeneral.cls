VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CargaAnalistas() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ERRORCargaAnalistas
    sSql = "select R.cPersCod, P.cPersNombre from RHEmpleado R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " where R.cRHEmpNivRem='007'"
    sSql = sSql & " order by P.cPersNombre"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CargaAnalistas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Function RecuperaParametro(Optional ByVal pnCodParametro As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaParametro
    sSql = "Select * from ColocParametro "
    If pnCodParametro <> -1 Then
        sSql = sSql & " WHERE nParamVar = " & pnCodParametro
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaParametro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametro:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaAgencias() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaAgencias
    sSql = "Select cAgeCod,cAgeDescripcion from Agencias"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencias:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
