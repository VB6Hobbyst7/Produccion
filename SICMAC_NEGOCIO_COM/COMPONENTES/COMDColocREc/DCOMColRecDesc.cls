VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColRecDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function mfgEstadoColocRecupDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColocEstRecVigJud
            lsDesc = "Vigente Judicial"
        Case gColocEstRecVigCast
            lsDesc = "Vigente Castigado"
        Case gColocEstRecCanJud
            lsDesc = "Cancelado"
        Case gColocEstRecCanCast
            lsDesc = "Cancelado"
    End Select
    mfgEstadoColocRecupDesc = lsDesc
End Function
