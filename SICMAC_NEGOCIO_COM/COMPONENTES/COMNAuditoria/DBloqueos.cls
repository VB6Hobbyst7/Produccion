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
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String

Public Function CargaBloqueo(Optional pnVarCod As CGBloqueos = -1) As Recordset
   On Error GoTo CargaBloqueoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion Then
      psSql = "SELECT nVarCod, cVarValor, cUltimaActualizacion FROM Bloqueos " & IIf(pnVarCod = -1, "", "WHERE nVarCod = " & pnVarCod)
      Set CargaBloqueo = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaBloqueoErr:
   MsgBox Err.Description, vbCritical
End Function

Public Sub ActualizaBloqueo(pnVarCod As CGBloqueos, psVarValor As String, psMovNro As String)
   On Error GoTo ActualizaBloqueoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion Then
      psSql = "UPDATE Bloqueos SET cVarValor = " & psVarValor & ", cUltimaActualizacion = '" & psMovNro & "' WHERE nVarCod = '" & pnVarCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaBloqueoErr:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
'    Dim oIni As New COMConecta.DCOMClasIni
'    Set oIni = Nothing
End Sub



