VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String

Public Function CargaBloqueo(Optional pnVarCod As CGBloqueos = -1) As Recordset
   On Error GoTo CargaBloqueoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "SELECT nVarCod, cVarValor, cUltimaActualizacion FROM Bloqueos " & IIf(pnVarCod = -1, "", "WHERE nVarCod = " & pnVarCod)
      Set CargaBloqueo = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaBloqueoErr:
   Call RaiseError(MyUnhandledError, "DBloqueos:CargaBloqueo Method")
End Function

Public Sub InsertaBloqueo(pnVarCod As CGBloqueos, psVarValor As String, psMovNro As String)
   On Error GoTo InsertaBloqueoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "INSERT Bloqueos (nVarCod, cVarValor, cUltimaActualizacion) " _
            & "VALUES (" & pnVarCod & ",'" & psVarValor & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaBloqueoErr:
   Call RaiseError(MyUnhandledError, "DBloqueos:InsertaBloqueo Method")
End Sub

Public Sub ActualizaBloqueo(pnVarCod As CGBloqueos, psVarValor As String, psMovNro As String)
   On Error GoTo ActualizaBloqueoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "UPDATE Bloqueos SET cVarValor = " & psVarValor & ", cUltimaActualizacion = '" & psMovNro & "' WHERE nVarCod = '" & pnVarCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaBloqueoErr:
   Call RaiseError(MyUnhandledError, "DBloqueos:ActualizaBloqueo Method")
End Sub

Public Sub EliminaBloqueos(pnVarCod As CGBloqueos)
   On Error GoTo EliminaBloqueosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "DELETE Bloqueos WHERE nVarCod = " & pnVarCod & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaBloqueosErr:
   Call RaiseError(MyUnhandledError, "DBloqueos:EliminaBloqueos Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
Set oIni = Nothing
End Sub



