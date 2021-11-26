VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosTablaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9701FC026C"
'Actualiza Tabla Alias
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902DA0110
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902DA01BA
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902DA019C
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega a una tabla o concepto un alias
'##ModelId=3A9B03F9036F
Public Function AgregaTablaAlias(pcCodTabCam As String, pcTablaColumna As String, pcAlias As String) As Boolean
    On Error GoTo AgregaTablaAliasErr

    'your code goes here...

    Exit Function
AgregaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "ActualizaDatosTablaCampo:AgregaTablaAlias Method")
End Function

'Elimina  alias de una tabla o concepto.
'##ModelId=3A9B03F90397
Public Function EliminaTablaAlias(pcCodTabCam As String) As Boolean
    On Error GoTo EliminaTablaAliasErr

    'your code goes here...

    Exit Function
EliminaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "ActualizaDatosTablaCampo:EliminaTablaAlias Method")
End Function

'Actualiza alias de una tabla o concepto.
'##ModelId=3A9B03F903C9
Public Function ModificaTablaAlias(pcCodTabCam As String, pcTablaColumna As String, pcAlias As String) As Boolean
    On Error GoTo ModificaTablaAliasErr

    'your code goes here...

    Exit Function
ModificaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "ActualizaDatosTablaCampo:ModificaTablaAlias Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902DA012E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
