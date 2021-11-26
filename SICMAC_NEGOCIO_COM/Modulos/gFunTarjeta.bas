Attribute VB_Name = "gFunTarjeta"
Public Function CargaComXRepoTarjeta(ByVal pnMoneda As Integer) As Double
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oConect As DConecta

    Set oConect = New DConecta
    oConect.AbreConexion
    
    sql = "Exec stp_sel_RecuperaTarifComReposicion " & pnMoneda
    Set rs = oConect.CargaRecordSet(sql)
    
    If Not rs.EOF Then
       CargaComXRepoTarjeta = rs!nValor
    Else
       CargaComXRepoTarjeta = 0#
    End If
    
    RSClose rs
    
    oConect.CierraConexion
    Set oConect = Nothing
    
End Function

Public Function ValidaTarjAnt(ByVal psPersCod As String) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oConect As DConecta

    Set oConect = New DConecta
    oConect.AbreConexion
    
    sql = "Exec stp_sel_ValidaTarjAnt " & psPersCod
    Set rs = oConect.CargaRecordSet(sql)
    
    If Not rs.EOF Then
       ValidaTarjAnt = True
    Else
       ValidaTarjAnt = False
    End If
    
    RSClose rs
    
    oConect.CierraConexion
    Set oConect = Nothing
End Function

