VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'##ModelId=3A9C06F10056
Public Function EliminaAF(psBSCod As String, psSerie As String, psPeriodo As String) As Boolean
    On Error GoTo EliminaAFErr
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    sql = "Delete BSActivoFijo Where nAnio = '" & psPeriodo & "' And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sql
    
    Set oCon = Nothing
    Exit Function
EliminaAFErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:AgregaAsistencia Method")
End Function

#If DebugMode Then
    '##ModelId=3AB90303010F
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

