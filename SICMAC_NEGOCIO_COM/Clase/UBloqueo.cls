VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
Public Sub ShowForm(oFormPadre As Object, Optional isModal As Boolean = False)
    oFormPadre.Enabled = False
    'frmLoad.Inicio oFormPadre, isModal
End Sub

Public Sub CloseForm(oFormPadre As Object)  'ByVal OwnerForm As Variant
    oFormPadre.Enabled = True
    'Unload frmLoad
End Sub
