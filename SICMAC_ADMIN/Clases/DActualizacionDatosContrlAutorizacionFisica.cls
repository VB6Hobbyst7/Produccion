VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizacionsCtrlAutFisica"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9C0D090262"
'Actualiza en Base Autorización Fisica
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB903050003
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90305014E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903050112
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega a una nueva autorizacion fisica
'##ModelId=3A9C0FBA03AD
Public Function AgregaAutorizacionFisica(psCodPers As String, psAutorizacionTpo As String, pdSolicitadoInicio As Date, pdSolicitadoFin As Date, pdProgramadoInicio As Date, pdProgramadoFin As Date, pdEjecutadoInicio As Date, pdEjecutadoFin As Date, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAutorizacionFisicaErr

    'your code goes here...

    Exit Function
AgregaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:AgregaAutorizacionFisica Method")
End Function

'Elimina una autorizacion fisica
'##ModelId=3A9C0FBB0015
Public Function EliminaAutorizacionFisica(psCodPers As String, psAutorizacionTpo As String, pdSolicitadoInicio As Date) As Boolean
    On Error GoTo EliminaAutorizacionFisicaErr

    'your code goes here...

    Exit Function
EliminaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:EliminaAutorizacionFisica Method")
End Function

'Actualiza una autorizacion fisica
'##ModelId=3A9C0FBB005B
Public Function ModificaAutorizacionFisica(psCodPers As String, psAutorizacionTpo As String, pdSolicitadoInicio As Date, pdSolicitadoFin As Date, pdProgramadoInicio As Date, pdProgramadoFin As Date, pdEjecutadoInicio As Date, pdEjecutadoFin As Date, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAutorizacionFisicaErr

    'your code goes here...

    Exit Function
ModificaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:ModificaAutorizacionFisica Method")
End Function

'Obtiene la autorizacion de descanso fisico, en un rango de fechas
'##ModelId=3A9C102D0204
Public Function GetAutorizacionFisica(psCodPers As String, pdInicio As Date, pdFin As Date) As Integer
    On Error GoTo GetAutorizacionFisicaErr

    'your code goes here...

    Exit Function
GetAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
End Function

#If DebugMode Then
    '##ModelId=3AB90305005E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
