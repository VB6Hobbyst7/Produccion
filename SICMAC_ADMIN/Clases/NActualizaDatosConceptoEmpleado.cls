VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosConEmpleado"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9B25930316"
'Actualiza Datos de la Relacion concepto Empleado
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902DD02B9
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902DD0363
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902DD034F
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega una nueva relación entre concepto y Empleado
'##ModelId=3A9B25AB01C6
Public Function AgregaConceptoEmpleado(psRHPersCod As String, psRHConceptoCod As String, psUltimaActualizacion As String, psRHRefDescripcion As String, pnRHConceptoValor As Double) As Boolean
    On Error GoTo AgregaConceptoEmpleadoErr

    'your code goes here...

    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

'Elimina  una nueva relacion entre concepto y Empleado
'
'##ModelId=3A9B25AB01F8
Public Function EliminaConceptoEmpleado(psRHPersCod As String, psRHConceptoCod As String) As Boolean
    On Error GoTo EliminaConceptoEmpleadoErr

    'your code goes here...

    Exit Function
EliminaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoEmpleado:EliminaConceptoEmpleado Method")
End Function

'Actualiza una nueva relacion entre concepto y Empleado
'##ModelId=3A9B25AB022A
Public Function ModificaConceptoEmpleado(psRHPersCod As String, psRHConceptoCod As String, psUltimaActualizacion As String, psRHRefDescripcion As String, pnRHConceptoValor As Double) As Boolean
    On Error GoTo ModificaConceptoEmpleadoErr

    'your code goes here...

    Exit Function
ModificaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoEmpleado:ModificaConceptoEmpleado Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902DD02D7
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
