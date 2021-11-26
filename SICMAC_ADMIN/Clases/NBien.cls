VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'** Nombre : NBien
'** Descripción : Para manejo de Bienes Logistica creado segun ERS059-2013
'** Creación : EJVG, 20130520 17:00:00 PM
'********************************************************************
Option Explicit

Public Function ListarOrdenCompra(ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim obj As New DBien
    Set ListarOrdenCompra = obj.ListarOrdenCompra(pnMoneda)
    Set obj = Nothing
End Function
