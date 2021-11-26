VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCapCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sCtaCod As String
Public sProducto As String
Public sMoneda As String
Public sRelacion As String
Public sEstado As String
Public sPersoneria As String


Public Sub CargaDatos(ByVal sCuenta As String, ByVal sProd As String, _
        ByVal sMon As String, ByVal sRel As String)
sCtaCod = sCuenta
sProducto = sProd
sMoneda = sMon
sRelacion = sRel
End Sub

Public Sub CargaDatos2(ByVal sCuenta As String, ByVal sProd As String, _
        ByVal sMon As String, ByVal sEst As String, ByVal sPers As String)
sCtaCod = sCuenta
sProducto = sProd
sMoneda = sMon
sEstado = sEst
sPersoneria = sPers


End Sub


