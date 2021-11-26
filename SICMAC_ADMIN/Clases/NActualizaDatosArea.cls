VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosArea"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A942128029A"
'Realiza la actualización de los datos del area
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D101D5
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D10276
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D10262
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

'Agrega una nuevo area Organizacional
'##ModelId=3A94216B0341
Public Function AgregaAreaOrg(psAreaCod As String, psAreaDescripcion As String, psLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAreaOrgErr

    'your code goes here...

    Exit Function
AgregaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosArea:AgregaAreaOrg Method")
End Function

'Elimina Area Organizacional, la eliminacion es logica
'##ModelId=3A94216B035F
Public Function EliminaAreaOrg(psAreaCod As String) As Boolean
    On Error GoTo EliminaAreaOrgErr

    'your code goes here...

    Exit Function
EliminaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosArea:EliminaAreaOrg Method")
End Function

'Actualiza un Area Organizacional
'##ModelId=3A94216B0387
Public Function ModificaAreaOrg(psAreaCod As String, psAreaDescripcion As String, psLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAreaOrgErr

    'your code goes here...

    Exit Function
ModificaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosArea:ModificaAreaOrg Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D101F3
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
