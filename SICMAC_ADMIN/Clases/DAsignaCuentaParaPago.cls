VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAsignaCuentaParaPago"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96C591029F"
'Asigna la cuenta para a la que se va a depositar el pago al RRHH de la Empresa
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902F80326
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902F9007E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F9004C
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

'Asigna una cuenta a un RRHH para que se le abone sus pagos
'##ModelId=3A96C7DB019E
Public Function AgregaRHCuenta(psPersCod As String, psCodTarj As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRHCuentaErr

    'your code goes here...

    Exit Function
AgregaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "DAsignaCuentaParaPago:AgregaRHCuenta Method")
End Function

'Elimina una cuenta Asignada a un RRHH.
'##ModelId=3A96C7DB01DA
Public Function EliminaRHCuenta(psPersCod As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo EliminaRHCuentaErr

    'your code goes here...

    Exit Function
EliminaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "DAsignaCuentaParaPago:EliminaRHCuenta Method")
End Function

'Asigna una cuenta a un RRHH para que se le abone sus pagos
'##ModelId=3A96C7DB020C
Public Function ModificaRHCuenta(psPersCod As String, psCodTarj As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRHCuentaErr

    'your code goes here...

    Exit Function
ModificaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "DAsignaCuentaParaPago:ModificaRHCuenta Method")
End Function

'##ModelId=3A96CC13010A
Public Function GetCuentas(cPersCod As String) As Recordset
    On Error GoTo GetCuentasErr

    'your code goes here...

    Exit Function
GetCuentasErr:
    Call RaiseError(MyUnhandledError, "DAsignaCuentaParaPago:GetCuentas Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902F80380
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
