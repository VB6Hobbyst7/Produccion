VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrevioFinan"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub Show(psCadImpresion As String, psTitulo As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    Dim o As clsConstImp
    oImpresora.Inicia pImpresora
    lTpoImpresora = pImpresora
    Inicia pImpresora
'    'By Capi 24102008
'    If pImpresora = gEPSON Then
'        pbCondensado = True
'    End If
'    '
    frmPrevio.PrevioFinan psCadImpresion, psTitulo, pbCondensado, pnLineas
End Sub

Public Sub ShowImpreSpool(ByVal psCadImpresion As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    Dim i As Long
    frmImpresora.Show 1
    If Not lbCancela Then
        For i = 1 To lnNumCopias
            ImpreBegin pbCondensado, pnLineas
                Print #ArcSal, ImpreCarEsp(psCadImpresion)
            ImpreEnd
        Next
    End If
End Sub

Public Sub PrintSpool(ByVal psLpt As String, ByVal psCadImpresion As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    sLpt = psLpt
    ImpreBegin pbCondensado, pnLineas
        Print #ArcSal, ImpreCarEsp(psCadImpresion)
    ImpreEnd
End Sub


