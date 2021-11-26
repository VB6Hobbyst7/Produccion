VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UColRecComisionSelecciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sCodAbog As String
Public nComisionTipo As Integer
Public nComisionCod As Integer
Public nComisionValor As Double

Public Sub CargaDatos(ByVal psCodAbog As String, ByVal pnTipoComision As Integer, _
        ByVal pnCodComision As Integer, ByVal pnValComision As Double)
sCodAbog = psCodAbog
nComisionTipo = pnTipoComision
nComisionCod = pnCodComision
nComisionValor = pnValComision
End Sub


