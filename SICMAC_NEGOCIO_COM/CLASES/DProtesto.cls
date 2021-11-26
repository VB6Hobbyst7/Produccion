VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DProtesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function ObtAgenProtesto() As ADODB.Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    On Error GoTo ErrHandler
        
        sSQL = "Select A.cAgeCod,A.cAgeDescripcion,AP.Monto"
        sSQL = sSQL & " From Agencias A"
        sSQL = sSQL & " Inner Join AgenProtesto AP on A.cAgeCod=AP.cCodAgen"
        
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set ObtAgenProtesto = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
        Set oConec = Nothing

    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ObtAgenProtesto = Null
End Function


Public Function ActualizaProtesto(ByVal pAgencia As String, ByVal pValor As Double) As Boolean
    Dim sSQL As String
    Dim oConec As DConecta
    On Error GoTo ErrHandler
    
        sSQL = "Update AgenProtesto set Monto=" & pValor & " Where cCodAgen='" & pAgencia & "'"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        oConec.ConexionActiva.BeginTrans
        oConec.ConexionActiva.Execute sSQL
        oConec.ConexionActiva.CommitTrans
        oConec.CierraConexion
        Set oConec = Nothing
        ActualizaProtesto = True
    Exit Function
ErrHandler:
        oConec.ConexionActiva.RollbackTrans
        If Not oConec Is Nothing Then Set oConec = Nothing
        ActualizaProtesto = False
End Function
