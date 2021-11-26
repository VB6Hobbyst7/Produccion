VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nGrupoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetGrupoOpe() As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta


If oConect.AbreConexion = False Then Exit Function

sql = " Select  cGrupoCod, cGrupoNombre, nOrden, cIngEgr,nEfectivo,cOPSuma,cConsDescripcion, nTipoGrupo " _
    & " From    GrupoOpe G " _
    & "         JOIN CONSTANTE ON G.nTipoGrupo = nConsValor AND  nConsCod=1031 ORDER BY cGrupoCod  "

Set rs = oConect.CargaRecordSet(sql)
Set GetGrupoOpe = rs
oConect.CierraConexion
Set oConect = Nothing

End Function
Public Function GetOpeGrupo(Optional ByVal psGrupo As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function
sql = " Select  O.cOpeCod, O.cOpeDesc " _
    & " From    Opetpo O JOIN GruposOpe GO ON GO.cOpeCod=O.cOpeCod " _
    & " Where   GO.cGrupoCod ='" & psGrupo & "'"

Set rs = oConect.CargaRecordSet(sql)
Set GetOpeGrupo = rs
oConect.CierraConexion
Set oConect = Nothing

End Function
Public Function GetOperaciones() As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function
sql = " Select  O.cOpeCod, O.cOpeDesc " _
    & " From    Opetpo o "
    

Set rs = oConect.CargaRecordSet(sql)
Set GetOperaciones = rs
oConect.CierraConexion
Set oConect = Nothing

End Function
Public Function GrabaOPeGrupo(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim oOpeGrup As DGrupoOpe
Set oOpeGrup = New DGrupoOpe
oOpeGrup.InsertaGruposOpe psOpeCod, psCodGrupoOpe
Set oOpeGrup = Nothing
End Function
Public Function EliminaGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim oOpeGrup As DGrupoOpe
Set oOpeGrup = New DGrupoOpe
oOpeGrup.DeleteGruposOpe psOpeCod, psCodGrupoOpe
Set oOpeGrup = Nothing
End Function
Public Function InsertaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer
Dim oOpeGrup As DGrupoOpe
Set oOpeGrup = New DGrupoOpe
InsertaOpeGrupo = oOpeGrup.InsertaOpeGrupo(psGrupoCod, psGrupoNombre, pnOrden, psIngEgr, pnEfectivo, pnTipoGrupo, psOPSuma)
Set oOpeGrup = Nothing
End Function
Public Function ActualizaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                    ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer
Dim oOpeGrup As DGrupoOpe
Set oOpeGrup = New DGrupoOpe
ActualizaOpeGrupo = oOpeGrup.ActOpeGrupo(psGrupoCod, psGrupoNombre, pnOrden, psIngEgr, pnEfectivo, pnTipoGrupo, psOPSuma)
Set oOpeGrup = Nothing
End Function

Public Function DeleteOpeGrupo(ByVal psGrupoCod As String) As Integer
Dim oOpeGrup As DGrupoOpe
Set oOpeGrup = New DGrupoOpe
DeleteOpeGrupo = oOpeGrup.DeleteOpeGrupo(psGrupoCod)
Set oOpeGrup = Nothing
End Function
Public Function VerificaGruposOpe(ByVal psGrupoCod As String, ByVal psOpeCod As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
VerificaGruposOpe = False
Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function
sql = " Select * From GruposOpe Where cGrupoCod='" & psGrupoCod & "' and cOpeCod='" & psOpeCod & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    VerificaGruposOpe = True
End If
oConect.CierraConexion
Set oConect = Nothing

End Function


