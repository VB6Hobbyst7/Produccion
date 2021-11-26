VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosRelaPers"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9454950053"
'Actualiza la relacion entre personas
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D3026E
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D3030F
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D302FB
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

'Agrega una nueva relacion entre personas
'##ModelId=3A9476C801E6
Public Function AgregaRelPers(psCodPers As String, psPersRelacCodPers As String, psPersRelac As Integer, pbPersRelacBenef As Boolean, pnPersRelacBenefPorc As Double, pbPersRelacAMP As Boolean, pbPersRelacEstado As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRelPersErr

    'your code goes here...

    Exit Function
AgregaRelPersErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRelacionesPers:AgregaRelPers Method")
End Function

'Elimina  una relacion entre personas
'##ModelId=3A9476C8020E
Public Function EliminaRelPers(psCodPers As String, psPersRelacCodPers As String) As Boolean
    On Error GoTo EliminaRelPersErr

    'your code goes here...

    Exit Function
EliminaRelPersErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRelacionesPers:EliminaRelPers Method")
End Function

'Actualiza relacion entre personas
'##ModelId=3A9476C80236
Public Function ModificaRelPers(psCodPers As String, psPersRelacCodPers As String, psPersRelac As Integer, pbPersRelacBenef As Boolean, pnPersRelacBenefPorc As Double, pbPersRelacAMP As Boolean, pbPersRelacEstado As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRelPersErr

    'your code goes here...

    Exit Function
ModificaRelPersErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRelacionesPers:ModificaRelPers Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D30282
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
