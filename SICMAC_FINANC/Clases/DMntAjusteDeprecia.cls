VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMntAjusteDeprecia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7F4CCF00CB"
Option Base 0
Option Explicit

'##ModelId=3A7F4D9C00EA
Public Sub CargaRubros()
   On Error GoTo CargaRubrosErr

    'your code goes here...

   Exit Sub
CargaRubrosErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:CargaRubros Method")
End Sub

'##ModelId=3A7F4F250271
Public Sub ObtieneNroItem()
   On Error GoTo ObtieneNroItemErr

    'your code goes here...

   Exit Sub
ObtieneNroItemErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:ObtieneNroItem Method")
End Sub

'##ModelId=3A7F4F3701A5
Public Sub InsertaRubro()
   On Error GoTo InsertaRubroErr

    'your code goes here...

   Exit Sub
InsertaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:InsertaRubro Method")
End Sub

'##ModelId=3A7F4F40006D
Public Sub ActualizaRubro()
   On Error GoTo ActualizaRubroErr

    'your code goes here...

   Exit Sub
ActualizaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:ActualizaRubro Method")
End Sub

'##ModelId=3A7F4F4302BF
Public Sub EliminaRubro()
   On Error GoTo EliminaRubroErr

    'your code goes here...

   Exit Sub
EliminaRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:EliminaRubro Method")
End Sub

'##ModelId=3A7F4EF402EE
Public Sub CargaDetRubros()
   On Error GoTo CargaDetRubrosErr

    'your code goes here...

   Exit Sub
CargaDetRubrosErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:CargaDetRubros Method")
End Sub

'##ModelId=3A7F4F460232
Public Sub InsertaDetRubro()
   On Error GoTo InsertaDetRubroErr

    'your code goes here...

   Exit Sub
InsertaDetRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:InsertaDetRubro Method")
End Sub

'##ModelId=3A7F4F6800EA
Public Sub ActualizaDetRubro()
   On Error GoTo ActualizaDetRubroErr

    'your code goes here...

   Exit Sub
ActualizaDetRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:ActualizaDetRubro Method")
End Sub

'##ModelId=3A7F4F6E01F4
Public Sub EliminaDetRubro()
   On Error GoTo EliminaDetRubroErr

    'your code goes here...

   Exit Sub
EliminaDetRubroErr:
   Call RaiseError(MyUnhandledError, "DMntAjusteDeprecia:EliminaDetRubro Method")
End Sub
