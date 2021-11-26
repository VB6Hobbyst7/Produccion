VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosProyeccionLab"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9D41F10023"
'Actualiza los datos de la proyección
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902E500D0
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902E5017A
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902E5015C
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

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9D5E370175
Public Function AgregaProyeccionLaboral(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo AgregaProyeccionLaboralErr

    'your code goes here...

    Exit Function
AgregaProyeccionLaboralErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosProyeccionLaboral:AgregaProyeccionLaboral Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9D5E370229
Public Function EliminaProyeccionLaboral(psRHPersCod As String, pdRHProyeccionFecha As Date) As Boolean
    On Error GoTo EliminaProyeccionLaboralErr

    'your code goes here...

    Exit Function
EliminaProyeccionLaboralErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosProyeccionLaboral:EliminaProyeccionLaboral Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9D5E37025B
Public Function ModificaProyeccionLaboral(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo ModificaProyeccionLaboralErr

    'your code goes here...

    Exit Function
ModificaProyeccionLaboralErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosProyeccionLaboral:ModificaProyeccionLaboral Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902E500EE
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
