VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NLogProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function ActDesProveedor(ByVal psPersCod As String, ByVal pnEstProv As Integer) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSql As String
    
    Set clsNProv = New DLogProveedor
    If pnEstProv = gLogProvEstadoActivado Then
        sSql = "UPDATE Proveedor SET cProvEstado = '" & gLogProvEstadoDesactivado & "' " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ActDesProveedor = clsNProv.ActDesProveedor(sSql)
    ElseIf pnEstProv = gLogProvEstadoDesactivado Then
        sSql = "UPDATE Proveedor SET cProvEstado = '" & gLogProvEstadoActivado & "' " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ActDesProveedor = clsNProv.ActDesProveedor(sSql)
    Else
        ActDesProveedor = 5
    End If
End Function
