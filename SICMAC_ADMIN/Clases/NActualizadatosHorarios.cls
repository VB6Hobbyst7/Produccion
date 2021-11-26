VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizadatosHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9C13EC0072"
'Actualiza Horarios
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902E302F4
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902E4001A
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902E303E4
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega a un nuevo horario tabla
'##ModelId=3A9C14210390
Public Function AgregaHorariosTabla(pcRHHorarioCod As String, pcRHHorarioDescripcion As String, pcUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaHorariosTablaErr

    'your code goes here...

    Exit Function
AgregaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:AgregaHorariosTabla Method")
End Function

'Elimina un horario tabla
'##ModelId=3A9C142103C2
Public Function EliminaHorariosTabla(pcRHHorarioCod As String) As Boolean
    On Error GoTo EliminaHorariosTablaErr

    'your code goes here...

    Exit Function
EliminaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:EliminaHorariosTabla Method")
End Function

'Actualiza un Horario tabla
'##ModelId=3A9C14220002
Public Function ModificaHorariosTabla(pcRHHorarioCod As String, pcRHHorarioDescripcion As String, pcUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaHorariosTablaErr

    'your code goes here...

    Exit Function
ModificaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:ModificaHorariosTabla Method")
End Function

'Asigna a un horario a una persona
'##ModelId=3A9C16BC0172
Public Function AgregaHorarios(pcRHPersCod As String, pdRHHorarioFecha As Date, pcUltimaActualizacion As String, pcRHHorarioCod As String, pcRHHorarioComentario As String) As Boolean
    On Error GoTo AgregaHorariosErr

    'your code goes here...

    Exit Function
AgregaHorariosErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:AgregaHorarios Method")
End Function

'Elimina una relacion horario persona
'##ModelId=3A9C16BC019A
Public Function EliminaHorarios(pcRHPersCod As String, pdRHHorarioFecha As Date) As Boolean
    On Error GoTo EliminaHorariosErr

    'your code goes here...

    Exit Function
EliminaHorariosErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:EliminaHorarios Method")
End Function

'Actualiza una relacion horario persona
'##ModelId=3A9C16BC01D6
Public Function ModificaHorarios(pcRHPersCod As String, pdRHHorarioFecha As Date, pcUltimaActualizacion As String, pcRHHorarioCod As String, pcRHHorarioComentario As String) As Boolean
    On Error GoTo ModificaHorariosErr

    'your code goes here...

    Exit Function
ModificaHorariosErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:ModificaHorarios Method")
End Function

'Agrega a un nuevo horario detalle
'##ModelId=3A9C175502C6
Public Function AgregaHorariosDetalle(pcRHHorarioCod As String, pcRHHorarioTurno As String, pcRHHorarioDias As String, pdRHHorarioInicio As Date, pdRHHorarioInicio As Date, pcUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaHorariosDetalleErr

    'your code goes here...

    Exit Function
AgregaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:AgregaHorariosDetalle Method")
End Function

'Elimina un horario detalle
'##ModelId=3A9C175502F8
Public Function EliminaHorariosDetalle(pcRHHorarioCod As String, pcRHHorarioTurno As String) As Boolean
    On Error GoTo EliminaHorariosDetalleErr

    'your code goes here...

    Exit Function
EliminaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:EliminaHorariosDetalle Method")
End Function

'Actualiza un Horario detalle
'##ModelId=3A9C17550321
Public Function ModificaHorariosDetalle(pcRHHorarioCod As String, pcRHHorarioTurno As String, pcRHHorarioDias As String, pdRHHorarioInicio As Date, pdRHHorarioInicio As Date, pcUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaHorariosDetalleErr

    'your code goes here...

    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "NActualizadatosHorarios:ModificaHorariosDetalle Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902E30312
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
