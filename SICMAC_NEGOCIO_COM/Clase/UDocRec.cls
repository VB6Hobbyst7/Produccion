VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UDocRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fnTpoDoc As Integer
Public fsNroDoc As String
Public fsPersCod As String
Public fsIFTpo As String
Public fsIFCta As String
Public fsPersNombre As String
Public fsDetalle As String
Public fnMonto As Currency
Public fsGlosa As String
Public fnNroCliLote As Long

Public Sub CargaDatos(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String, _
                        ByVal psPersNombre As String, ByVal psDetalle As String, ByVal pnMonto As Currency, ByVal psGlosa As String)
    fnTpoDoc = pnTpoDoc
    fsNroDoc = psNroDoc
    fsPersCod = psPersCod
    fsIFTpo = psIFTpo
    fsIFCta = psIFCta
    fsPersNombre = psPersNombre
    fsDetalle = psDetalle
    fnMonto = pnMonto
    fsGlosa = psGlosa
    fnNroCliLote = val(Left(fsDetalle, 10))
End Sub


