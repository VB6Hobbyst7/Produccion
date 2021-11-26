VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCapCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sCtaCod As String
Public sProducto As String
Public sMoneda As String
Public sRelacion As String
'FRHU 20140522 ERS068-2014 RQ14283
Public sNumSolictud As String
Public sPersNombre As String
Public dFechaSolicitud As String
Public sMonto As String
Public sCuenta As String
'FIN FRHU 20140522

Public Sub CargaDatos(ByVal sCuenta As String, ByVal sProd As String, _
        ByVal sMon As String, ByVal sRel As String)
sCtaCod = sCuenta
sProducto = sProd
sMoneda = sMon
sRelacion = sRel
End Sub

'FRHU 20140522 ERS068-2014 RQ14283
Public Sub CargaDatosSegTarj(ByVal psNumSolictud As String, ByVal psPersNombre As String, _
        ByVal psCuenta As String, ByVal psMonto As String)
sNumSolictud = psNumSolictud
sPersNombre = psPersNombre
sCuenta = psCuenta
sMonto = psMonto
sCuenta = psCuenta
End Sub
'FIN FRHU 20145022

