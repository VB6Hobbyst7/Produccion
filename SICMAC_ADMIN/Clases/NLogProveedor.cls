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
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor
    If pnEstProv = gLogProvEstadoActivado Then
        sSQL = "UPDATE Proveedor SET nProvEstado = " & gLogProvEstadoDesactivado & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ActDesProveedor = clsNProv.ActDesProveedor(sSQL)
    ElseIf pnEstProv = gLogProvEstadoDesactivado Then
        sSQL = "UPDATE Proveedor SET nProvEstado = " & gLogProvEstadoActivado & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ActDesProveedor = clsNProv.ActDesProveedor(sSQL)
    Else
        ActDesProveedor = 5
    End If
End Function

Public Function ProveedorActAgeReten(ByVal psPersCod As String, ByVal pnEstProv As LogProvAgenteRetencion) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor
    If pnEstProv = gLogProvAgenteRetencionSi Then
        sSQL = "UPDATE Proveedor SET bAgenteRetencion = " & gLogProvAgenteRetencionNo & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ProveedorActAgeReten = clsNProv.ActDesProveedor(sSQL)
    ElseIf pnEstProv = gLogProvAgenteRetencionNo Then
        sSQL = "UPDATE Proveedor SET bAgenteRetencion = " & gLogProvAgenteRetencionSi & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ProveedorActAgeReten = clsNProv.ActDesProveedor(sSQL)
    Else
        ProveedorActAgeReten = 5
    End If
End Function

Public Function ProveedorBuenCont(ByVal psPersCod As String, ByVal pnEstProv As LogProvBuenContribuyente) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor
    If pnEstProv = gLogProvBuenContribuyenteSi Then
        sSQL = "UPDATE Proveedor SET bBuenContribuyente = " & gLogProvBuenContribuyenteNo & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ProveedorBuenCont = clsNProv.ActDesProveedor(sSQL)
    ElseIf pnEstProv = gLogProvBuenContribuyenteNo Then
        sSQL = "UPDATE Proveedor SET bBuenContribuyente = " & gLogProvBuenContribuyenteSi & " " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        ProveedorBuenCont = clsNProv.ActDesProveedor(sSQL)
    Else
        ProveedorBuenCont = 5
    End If
End Function

Public Function HabDesabilitaSunat(ByVal psPersCod As String, ByVal pnEstSunat As Integer) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor
    If pnEstSunat = 1 Then
        sSQL = "UPDATE Proveedor SET bhabilSUNAT = 0 " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        HabDesabilitaSunat = clsNProv.HabDesSunat(sSQL)
    ElseIf pnEstSunat = 0 Then
        sSQL = "UPDATE Proveedor SET bhabilSUNAT = 1 " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        HabDesabilitaSunat = clsNProv.HabDesSunat(sSQL)
    Else
        HabDesabilitaSunat = 5
    End If
End Function



Public Function IngresaQuitaComentario(ByVal psPersCod As String, ByVal sComentario As String) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor

        sSQL = "UPDATE Proveedor SET cComentario = '" & sComentario & "' " & _
            " WHERE cPersCod = '" & psPersCod & "'"
            
        IngresaQuitaComentario = clsNProv.IngresaQuitaComent(sSQL)
End Function


Public Function HabDesabilitaConsucode(ByVal psPersCod As String, ByVal pnEstConsucode As Integer) As Integer
    Dim clsNProv As DLogProveedor
    Dim sSQL As String
    
    Set clsNProv = New DLogProveedor
    If pnEstConsucode = 1 Then
        sSQL = "UPDATE Proveedor SET bhabilSUNAT = 0 " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        HabDesabilitaConsucode = clsNProv.HabDesConsucode(sSQL)
    ElseIf pnEstConsucode = 0 Then
        sSQL = "UPDATE Proveedor SET bhabilSUNAT = 1 " & _
            " WHERE cPersCod = '" & psPersCod & "'"
        HabDesabilitaConsucode = clsNProv.HabDesConsucode(sSQL)
    Else
        HabDesabilitaConsucode = 5
    End If
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub
