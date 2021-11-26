VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapacitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Sub InsertaRHCapacitacionAprobaciones(ByVal pnProgNro As Integer)

Dim sSQL As String
Dim oCapacitacion As New DCapacitacion
Dim rs As New adodb.Recordset

Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionAprobaciones
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    sSQL = "SELECT cRHCargoCod,nOrden FROM RHCapacitacionNiveles"
    Set rs = oConecta.CargaRecordSet(sSQL)
    
    While Not rs.EOF
        Call oCapacitacion.InsertaRHCapacitacionAprobacion(pnProgNro, rs("cRHCargoCod"), rs("nOrden"))
        rs.MoveNext
    Wend
    
    oConecta.CierraConexion

    Set oConecta = Nothing
    Set oCapacitacion = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionAprobaciones:
    Err.Raise Err.Number, "InsertaRHCapacitacionAprobaciones", Err.Description
    Set oConecta = Nothing
End Sub
