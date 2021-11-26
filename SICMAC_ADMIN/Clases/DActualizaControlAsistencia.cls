VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaControlAsistencia"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9C06E00369"
'Actualiza la asistencia y tardanzas del RRHH de la Empresa
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030300BF
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90303024F
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903030213
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
'##ModelId=3A9C06F10056
Public Function AgregaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAsistenciaErr

    'your code goes here...

    Exit Function
AgregaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:AgregaAsistencia Method")
End Function

'Elimina un ingreso o salida del RRHH
'##ModelId=3A9C06F1009C
Public Function EliminaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String) As Boolean
    On Error GoTo EliminaAsistenciaErr

    'your code goes here...

    Exit Function
EliminaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:EliminaAsistencia Method")
End Function

'Actualiza Elimina un ingreso o salida del RRHH
'##ModelId=3A9C06F100CE
Public Function ModificaAsistencia(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAsistenciaErr

    'your code goes here...

    Exit Function
ModificaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:ModificaAsistencia Method")
End Function

'Agrega un nuevo ingreso o salida del RRHH
'##ModelId=3A9C06F1010A
Public Function AgregaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
AgregaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:AgregaAsistenciaIngSal Method")
End Function

'Elimina un ingreso o salida del RRHH
'##ModelId=3A9C06F10150
Public Function EliminaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String) As Boolean
    On Error GoTo EliminaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
EliminaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:EliminaAsistenciaIngSal Method")
End Function

'Actualiza Elimina un ingreso o salida del RRHH
'##ModelId=3A9C06F1018C
Public Function ModificaAsistenciaIngSal(psRHPersCod As String, psRHAsistenciaFechaRef As String, pnRHAsistenciaTardanzasMin As Integer, pnRHAsistenciaPermisoMin As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAsistenciaIngSalErr

    'your code goes here...

    Exit Function
ModificaAsistenciaIngSalErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:ModificaAsistenciaIngSal Method")
End Function

'Calcula la asistencia real en un rango de fechas
'##ModelId=3A9C06F101C9
Public Function CalculaAsistencia(pdInicio As Date, pdFin As Date) As Boolean
    On Error GoTo CalculaAsistenciaErr

    'your code goes here...

    Exit Function
CalculaAsistenciaErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:CalculaAsistencia Method")
End Function

'Devuelve la asistencia de un rango de fecahs de un RRHH
'##ModelId=3A9C071802FB
Public Function GetAsistencia(psRRHHCod As String) As Recordset
    On Error GoTo GetAsistenciaErr

    'your code goes here...

    Exit Function
GetAsistenciaErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:GetAsistencia Method")
End Function

#If DebugMode Then
    '##ModelId=3AB90303010F
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
