VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosTablaAlias"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9B05380024"
'Actualiza Tabla Alias
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FD039B
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902FE0108
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902FE00C2
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

'Agrega a una tabla o concepto un alias
'##ModelId=3A9B06C903C4
Public Function AgregaTablaAlias(psCodTabCam As String, psTablaColumna As String, psAlias As String) As Boolean
    On Error GoTo AgregaTablaAliasErr

    'your code goes here...

    Exit Function
AgregaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosTablaAlias:AgregaTablaAlias Method")
End Function

'Elimina  alias de una tabla o concepto.
'##ModelId=3A9B06CA0018
Public Function EliminaTablaAlias(psCodTabCam As String) As Boolean
    On Error GoTo EliminaTablaAliasErr

    'your code goes here...

    Exit Function
EliminaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosTablaAlias:EliminaTablaAlias Method")
End Function

'Actualiza alias de una tabla o concepto.
'##ModelId=3A9B06CA0054
Public Function ModificaTablaAlias(psCodTabCam As String, psTablaColumna As String, psAlias As String) As Boolean
    On Error GoTo ModificaTablaAliasErr

    'your code goes here...

    Exit Function
ModificaTablaAliasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosTablaAlias:ModificaTablaAlias Method")
End Function

'Devuelve la tabla alias
'##ModelId=3A9B07000341
Public Function GetTablaAlias() As Recordset
    On Error GoTo GetTablaAliasErr

    'your code goes here...

    Exit Function
GetTablaAliasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosTablaAlias:GetTablaAlias Method")
End Function

'Devuelve nombre de la tabla  a partir de un alias
'##ModelId=3A9B0713037A
Public Function GetNombreTabla(psAlias As String) As String
    On Error GoTo GetNombreTablaErr

    'your code goes here...

    Exit Function
GetNombreTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosTablaAlias:GetNombreTabla Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902FE000D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
