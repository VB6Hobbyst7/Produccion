VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMntBienesAdjudica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7F56B10232"
Option Base 0
Option Explicit

'##ModelId=3A7F571E0280
Public Sub CargaRubros()
   On Error GoTo CargaRubrosErr

    'your code goes here...

   Exit Sub
CargaRubrosErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:CargaRubros Method")
End Sub

'##ModelId=3A7F56C000FA
Public Sub InsertaRubro()
   On Error GoTo InsertaRubroErr

    'your code goes here...

   Exit Sub
InsertaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:InsertaRubro Method")
End Sub

'##ModelId=3A7F56C6029F
Public Sub ActualizaRubro()
   On Error GoTo ActualizaRubroErr

    'your code goes here...

   Exit Sub
ActualizaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:ActualizaRubro Method")
End Sub

'##ModelId=3A7F56CD005D
Public Sub EliminaRubro()
   On Error GoTo EliminaRubroErr

    'your code goes here...

   Exit Sub
EliminaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:EliminaRubro Method")
End Sub

'##ModelId=3A7F57250119
Public Sub CargaBienes()
   On Error GoTo CargaBienesErr

    'your code goes here...

   Exit Sub
CargaBienesErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:CargaBienes Method")
End Sub

'##ModelId=3A7F572A0157
Public Sub InsertaBienes()
   On Error GoTo InsertaBienesErr

    'your code goes here...

   Exit Sub
InsertaBienesErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:InsertaBienes Method")
End Sub

'##ModelId=3A7F5738033C
Public Sub ActualizaBienes()
   On Error GoTo ActualizaBienesErr

    'your code goes here...

   Exit Sub
ActualizaBienesErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:ActualizaBienes Method")
End Sub

'##ModelId=3A7F573D01A5
Public Sub EliminaBienes()
   On Error GoTo EliminaBienesErr

    'your code goes here...

   Exit Sub
EliminaBienesErr:
   Call RaiseError(MyUnhandledError, "DMntBienesAdjudica:EliminaBienes Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub
