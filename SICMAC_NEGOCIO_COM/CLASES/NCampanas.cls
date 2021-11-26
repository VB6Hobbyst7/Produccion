VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GeneraIdCampanas() As Integer
    Dim nIdCampana As Integer
    Dim oDCampanas As DCampanas
    
    Set oDCampanas = New DCampanas
    nIdCampana = oDCampanas.ObtenerMaxIdCampanas
    Set oDCampanas = Nothing
    
    nIdCampana = nIdCampana + 1
    GeneraIdCampanas = nIdCampana
End Function
