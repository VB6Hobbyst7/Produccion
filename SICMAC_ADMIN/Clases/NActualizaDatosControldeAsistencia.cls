VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosCtrAsist"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9BFEA602D4"
'Actualiza la asistencia y tardanzas del RRHH de la Empresa
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902E00231
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902E00335
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902E00321
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

'Agrega un nuevo ingreso o salida del RRHH
'##ModelId=3A9C01830256
Public Function AgregaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAsistenciaErr

    'your code goes here...

    Exit Function
AgregaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:AgregaAsistencia Method")
End Function

'Elimina un ingreso o salida del RRHH
'##ModelId=3A9C01830288
Public Function EliminaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String) As Boolean
    On Error GoTo EliminaAsistenciaErr

    'your code goes here...

    Exit Function
EliminaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:EliminaAsistencia Method")
End Function

'Actualiza Elimina un ingreso o salida del RRHH
'##ModelId=3A9C018302BA
Public Function ModificaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAsistenciaErr

    'your code goes here...

    Exit Function
ModificaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:ModificaAsistencia Method")
End Function

'Agrega un nuevo ingreso o salida del RRHH
'##ModelId=3A9C063C0150
Public Function AgregaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
AgregaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:AgregaAsistenciaIngSal Method")
End Function

'Elimina un ingreso o salida del RRHH
'##ModelId=3A9C063C0182
Public Function EliminaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String) As Boolean
    On Error GoTo EliminaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
EliminaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:EliminaAsistenciaIngSal Method")
End Function

'Actualiza Elimina un ingreso o salida del RRHH
'##ModelId=3A9C063C01B4
Public Function ModificaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
ModificaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:ModificaAsistenciaIngSal Method")
End Function

'Calcula la asistencia real en un rango de fechas
'##ModelId=3A9C065C0228
Public Sub CalculaAsistencia(pdInicio As Date, pdFin As Date)
    On Error GoTo CalculaAsistenciaErr

    'your code goes here...

    Exit Sub
CalculaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosControldeAsistencia:CalculaAsistencia Method")
End Sub

#If DebugMode Then
    '##ModelId=3AB902E00259
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
