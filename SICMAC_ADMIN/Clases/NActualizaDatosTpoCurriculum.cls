VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosTpoCurriculum"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A94185A0090"
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D00152
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D00260
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D0021A
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

'Agrega un Nuevo tipo de curriculum
'##ModelId=3A9418EB01F7
Public Function AgregaTpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaTpoCurriculumErr

    'your code goes here...

    Exit Function
AgregaTpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosTpoCurriculum:AgregaTpoCurriculum Method")
End Function

'Elimina un tipo de curriculum
'##ModelId=3A9418EB0229
Public Function EliminaITpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, nRHCurriculumItem As Integer, cUltimaActualizacion As String) As Boolean
    On Error GoTo EliminaITpoCurriculumErr

    'your code goes here...

    Exit Function
EliminaITpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosTpoCurriculum:EliminaITpoCurriculum Method")
End Function

'Actualiza un tipo de curriculum
'##ModelId=3A9418EB025B
Public Function ModificaITpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, nRHCurriculumItem As Integer, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaITpoCurriculumErr

    'your code goes here...

    Exit Function
ModificaITpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosTpoCurriculum:ModificaITpoCurriculum Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D00184
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
