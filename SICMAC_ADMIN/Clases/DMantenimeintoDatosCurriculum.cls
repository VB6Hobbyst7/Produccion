VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMantenimeintoDatosCurr"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9411E1039B"
'Realiza la actualizacion de los curriculums  y de los tipos de curriculums
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902E70368
    Private mlClassDebugID As Long
#End If

'Elimina el Item del curriculum de la persona indicada.
'##ModelId=3A94120E00CE
Public Function EliminaItemCurriculum(psRHPersCod As String, psRHCurriculumCod As String, nRHCurriculumItem As Integer, cUltimaActualizacion As String) As Boolean
    On Error GoTo EliminaItemCurriculumErr

    'your code goes here...

    Exit Function
EliminaItemCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:EliminaItemCurriculum Method")
End Function

'##ModelId=3AB902E80110
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902E800D4
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

'Agrega un Nuevo Item al curriculum de la persona indicada.
'##ModelId=3A94120E009C
Public Function AgregaItemCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaItemCurriculumErr

    'your code goes here...

    Exit Function
AgregaItemCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:AgregaItemCurriculum Method")
End Function

'Actualiza un item de un curriculum
'##ModelId=3A94175F0158
Public Function ModificaItemCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, nRHCurriculumItem As Integer, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaItemCurriculumErr

    'your code goes here...

    Exit Function
ModificaItemCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:ModificaItemCurriculum Method")
End Function

'Obtene un curriculum de la persona indicada
'##ModelId=3A9412450312
Public Function GetCurriculum(psRHPersCod As String) As Recordset
    On Error GoTo GetCurriculumErr

    'your code goes here...

    Exit Function
GetCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:GetCurriculum Method")
End Function

'Agrega un Nuevo tipo de curriculum
'##ModelId=3A941950004E
Public Function AgregaTpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaTpoCurriculumErr

    'your code goes here...

    Exit Function
AgregaTpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:AgregaTpoCurriculum Method")
End Function

'Elimina un tipo de curriculum
'##ModelId=3A941950008A
Public Function EliminaITpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, nRHCurriculumItem As Integer, cUltimaActualizacion As String) As Boolean
    On Error GoTo EliminaITpoCurriculumErr

    'your code goes here...

    Exit Function
EliminaITpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:EliminaITpoCurriculum Method")
End Function

'Actualiza un tipo de curriculum
'##ModelId=3A94195000C6
Public Function ModificaITpoCurriculum(psRHPersCod As String, psRHCurriculumCod As String, psRHCurriculumLugar As String, nRHCurriculumItem As Integer, pdRHCurriculumInicio As Date, dRHCurriculumFin As Date, psRHCurriculumUndTiempo As String, pnRHCurriculumTiempo As Integer, psRHCurriculumComentario As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaITpoCurriculumErr

    'your code goes here...

    Exit Function
ModificaITpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:ModificaITpoCurriculum Method")
End Function

'Obtene los tipod de curriculums
'##ModelId=3A9419550164
Public Function GetTpoCurriculum() As Recordset
    On Error GoTo GetTpoCurriculumErr

    'your code goes here...

    Exit Function
GetTpoCurriculumErr:
    Call RaiseError(MyUnhandledError, "DMantenimeintoDatosCurriculum:GetTpoCurriculum Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902E703CC
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
