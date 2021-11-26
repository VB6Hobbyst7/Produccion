VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ActualizaZonas(ByVal psCodigo As String, ByVal psDescripcion As String, ByVal pnAccion As Integer)
Dim oBase As DCredActualizaBD
    Set oBase = New DCredActualizaBD
    If pnAccion = 1 Then
        Call oBase.dInsertZonas(psCodigo, psDescripcion, False)
    Else
        If pnAccion = 2 Then
            Call oBase.dUpdateZonas(psCodigo, psDescripcion, False)
        Else
            Call oBase.dDeleteZonas(psCodigo, False)
        End If
    End If
    Set oBase = Nothing
End Sub

