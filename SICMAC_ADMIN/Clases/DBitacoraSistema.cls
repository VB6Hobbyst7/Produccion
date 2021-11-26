VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBitacoraSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oConect As DConecta

Private Sub Class_Initialize()
    Set oConect = New DConecta
    oConect.AbreConexion
End Sub

Public Sub BitarcoraSistemaInserta(pdFecha As Date, psUser As String, pnSistema As Integer, psAgeCod As String, psAreaCod As String, psPC As String, psProceso As String, Optional psComentario As String = "")
    Dim sql As String
    
    sql = "BitacoraSistemaInsert '" & Format(pdFecha, gsFormatoFechaHora) & "','" & psUser & "'," & pnSistema & ", '" & psAgeCod & "', '" & psAreaCod & "', '" & psPC & "','" & psProceso & "', '" & psComentario & "'"
    oConect.Ejecutar sql
End Sub

Private Sub Class_Terminate()
    oConect.CierraConexion
    Set oConect = Nothing
End Sub
