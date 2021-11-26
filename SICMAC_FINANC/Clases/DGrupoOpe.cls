VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGrupoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim oConect As DConecta

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oConect = New DConecta
Set oIni = New ClasIni
oConect.AbreConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function InsertaGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim sql As String
On Error GoTo InsertaGruposOpeErr
InsertaGruposOpe = 1
sql = "INSERT INTO GruposOpe(cGrupoCod, cOpeCod) Values('" & psCodGrupoOpe & "','" & psOpeCod & "')"
oConect.Ejecutar sql
InsertaGruposOpe = 0
Exit Function
InsertaGruposOpeErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function DeleteGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim sql As String
On Error GoTo DeleteGruposOpeErr
DeleteGruposOpe = 1
sql = "DELETE GruposOpe WHERE cGrupoCod='" & psCodGrupoOpe & "' AND cOpeCod='" & psOpeCod & "'"
oConect.Ejecutar sql
DeleteGruposOpe = 0
Exit Function
DeleteGruposOpeErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Class_Terminate()
oConect.AbreConexion
Set oConect = Nothing
End Sub
Public Function InsertaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer

Dim sql As String
On Error GoTo InsertaOpeGrupoErr
InsertaOpeGrupo = 1
sql = " INSERT INTO GrupoOpe(cGrupoCod, cGrupoNombre, nOrden, cIngEgr, nEfectivo, nTipoGrupo, cOPSuma) " _
    & " VALUES('" & psGrupoCod & "','" & Replace(psGrupoNombre, "'", "''") & "'," & pnOrden & ",'" & psIngEgr & "'," _
    & pnEfectivo & "," & pnTipoGrupo & ",'" & psOPSuma & "')"

oConect.Ejecutar sql
InsertaOpeGrupo = 0
Exit Function
InsertaOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function ActOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer

Dim sql As String
On Error GoTo ActOpeGrupoErr
ActOpeGrupo = 1

sql = " UPDATE  GrupoOpe " _
    & " SET     cGrupoNombre='" & Replace(psGrupoNombre, "'", "''") & "', " _
    & "         nOrden=" & pnOrden & "," _
    & "         cIngEgr='" & psIngEgr & "'," _
    & "         nEfectivo=" & pnEfectivo & "," _
    & "         nTipoGrupo=" & pnTipoGrupo & "," _
    & "         cOPSuma= '" & psOPSuma & "' " _
    & " WHERE  cGrupoCod='" & psGrupoCod & "'"

oConect.Ejecutar sql
ActOpeGrupo = 0
Exit Function
ActOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function DeleteOpeGrupo(ByVal psGrupoCod As String) As Integer
Dim sql As String
On Error GoTo DeleteOpeGrupoErr
DeleteOpeGrupo = 1
sql = " Delete GrupoOpe  WHERE  cGrupoCod='" & psGrupoCod & "'"
oConect.Ejecutar sql
DeleteOpeGrupo = 0
Exit Function
DeleteOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

