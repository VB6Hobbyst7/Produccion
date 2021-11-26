VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizadatosConcepto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9AFE2D0086"
'Actualiza Conceptos
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FC0097
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902FC01F5
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902FC01AF
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega un nuevo concepto
'##ModelId=3A9AFE7700D2
Public Function AgregaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As Boolean, pbRHConceptoMesTrab As Boolean, psRHConceptoCtaCont As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaConceptoErr

    'your code goes here...

    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:AgregaConcepto Method")
End Function

'Elimina  un concepto
'##ModelId=3A9AFE77010E
Public Function EliminaConcepto(psRHConceptoCod As String) As Boolean
    On Error GoTo EliminaConceptoErr

    'your code goes here...

    Exit Function
EliminaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:EliminaConcepto Method")
End Function

'Actualiza un concepto
'##ModelId=3A9AFE77014A
Public Function ModificaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As Boolean, pbRHConceptoMesTrab As Boolean, psRHConceptoCtaCont As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaConceptoErr

    'your code goes here...

    Exit Function
ModificaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:ModificaConcepto Method")
End Function

'Devulve los conceptos Activos,
'
'El parametro psTpoConcepto si es vacio dretorna todos los tipos de conceptos
'caso contrario retorna todos los conceptos del tipo indicado
'##ModelId=3A9AFE7A0380
Public Function GetConceptos(Optional psTpoConcepto As String = "") As Recordset
    On Error GoTo GetConceptosErr

    'your code goes here...

    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Devuelve la formula usada
'##ModelId=3A9AFE980347
Public Function GetFormula() As String
    On Error GoTo GetFormulaErr

    'your code goes here...

    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902FC00F1
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
