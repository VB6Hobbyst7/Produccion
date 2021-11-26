VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosRelaPers"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A947A890148"
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902EE02BD
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902EF0020
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902EE03C2
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
'##ModelId=3A9481E603E7
Public Function AgregaRelPers(psCodPers As String, psPersRelacCodPers As String, psPersRelac As Integer, pbPersRelacBenef As Boolean, pnPersRelacBenefPorc As Double, pbPersRelacAMP As Boolean, pbPersRelacEstado As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRelPersErr

    'your code goes here...

    Exit Function
AgregaRelPersErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRelacionesPers:AgregaRelPers Method")
End Function

'Elimina  una relacion entre personas
'##ModelId=3A9481E70031
Public Function EliminaRelPers(psCodPers As String, psPersRelacCodPers As String) As Boolean
    On Error GoTo EliminaRelPersErr

    'your code goes here...

    Exit Function
EliminaRelPersErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRelacionesPers:EliminaRelPers Method")
End Function

'Actualiza relacion entre personas
'##ModelId=3A9481E7006E
Public Function ModificaRelPers(psCodPers As String, psPersRelacCodPers As String, psPersRelac As Integer, pbPersRelacBenef As Boolean, pnPersRelacBenefPorc As Double, pbPersRelacAMP As Boolean, pbPersRelacEstado As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRelPersErr

    'your code goes here...

    Exit Function
ModificaRelPersErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRelacionesPers:ModificaRelPers Method")
End Function

'##ModelId=3A9481F4035B
Public Function GetRelPers(psCodPers As String) As Recordset
    On Error GoTo GetRelPersErr

    'your code goes here...

    Exit Function
GetRelPersErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRelacionesPers:GetRelPers Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902EE0317
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
