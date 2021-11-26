VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaAdeudados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A89DACF0109"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit

'emite una lista con todos los adeudados pendientes para las operaciones
'respectivas.
'##ModelId=3A89DC4A0014
Public Sub ListaAdeudPend(psCodCtaCont As String)
    On Error GoTo ListaAdeudPendErr

    'your code goes here...

    Exit Sub
ListaAdeudPendErr:
    Call RaiseError(MyUnhandledError, "NCajaAdeudados:ListaAdeudPend Method")
End Sub

'Calcula el Interes de ctas de aduedados los cual posee una particularidad con
'algunas ctas las cuales utilizan una Tasa de Interes VAC
'##ModelId=3A8ADEBE0007
Public Sub CalculaIntAdued(pnTasaVac As Currency, pnCapital As Currency, pnDias As Integer, pnPeriodo As Integer, pnTasaInt As Currency, psMonPag As String, psCodIFCta As String)
    On Error GoTo CalculaIntAduedErr

    'your code goes here...

    Exit Sub
CalculaIntAduedErr:
    Call RaiseError(MyUnhandledError, "NCajaAdeudados:CalculaIntAdued Method")
End Sub
Public Function CalculaInteres(lnDiasTrans As Long, lnPeriodo As Long, lnTasaInt As Currency, lnCapital As Currency) As Currency
    Dim lnFactor As Double
    Dim lnBaseInt As Double
    If lnPeriodo > 0 Then
        lnFactor = lnDiasTrans / lnPeriodo
        lnBaseInt = ((1 + (lnTasaInt / 100)) ^ lnFactor) - 1
        CalculaInteres = lnBaseInt * lnCapital
    Else
        CalculaInteres = 0
    End If
End Function
'##ModelId=3A931BCE00AC
Public Sub ReporteGenAdeud(psBanco As String, psOpeCod As String, pdFecha As Date)
    On Error GoTo ReporteGenAdeudErr

    'your code goes here...

    Exit Sub
ReporteGenAdeudErr:
    Call RaiseError(MyUnhandledError, "NCajaAdeudados:ReporteGenAdeud Method")
End Sub

'##ModelId=3A931F7200CA
Public Function ReporteDetAdeud(psCtaAdeudIni As String, psCtaAdeudFin As String, psCtaContCap As String, psCtaContInt As String, psOpeCod As String) As String
    On Error GoTo ReporteDetAdeudErr

    'your code goes here...

    Exit Function
ReporteDetAdeudErr:
    Call RaiseError(MyUnhandledError, "NCajaAdeudados:ReporteDetAdeud Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
