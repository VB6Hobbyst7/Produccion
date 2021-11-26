VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NAperturarCuentaRRHH"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A945476018F"
'Realiza la Apertura de cuenta de los empleados, la apertura tiene valores ya
'definidos, y tiene menos validaciones que una cuenta para clientes.
'Además se asigna automaticamente se relaciona estas cuenta como cuenta de
'deposito para el pago del RRHH.
'Siempre se valida si la perosna tinene cuentas en la agencia.
'Sepuede aperturar cuentas de RRHH en cualquier Agencia.
'Se deben imprimir todos los documentos de una apertura a un cliente.
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D20263
    Private mlClassDebugID As Long
#End If

'##ModelId=3A9475E30135
Public theNCaptaImpresion As NCaptaImpresion

'Apertura de Cuenta de cuentas de captaciones.
'
'Como parametro se pasan:
'
'   psMovNro : Numero Movimiento
'   pnMontoNotaABono : Monto de Nota de Abono
'   pClientes : recordset Clientes y Relaciones
'
'
'##ModelId=3A947574038C
Public Function AperturaCuentaRRHH(pdSistema As Date, psMovNro As String, pnMontoNotaAbono As Currency, psClientes As Recordset, Optional pbOrdenPago As Boolean = False, Optional psInstitucion As String, Optional psCodOpe As String = "") As Boolean
    On Error GoTo AperturaCuentaRRHHErr

    'your code goes here...

    Exit Function
AperturaCuentaRRHHErr:
    Call RaiseError(MyUnhandledError, "NAperturarCuentaRRHH:AperturaCuentaRRHH Method")
End Function

'##ModelId=3AB902D20321
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D20303
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
'##ModelId=3A96CBE5038F
Public Function AgregaRHCuenta(psPersCod As String, psCodTarj As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRHCuentaErr

    'your code goes here...

    Exit Function
AgregaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "NAperturarCuentaRRHH:AgregaRHCuenta Method")
End Function

'Elimina una cuenta Asignada a un RRHH.
'##ModelId=3A96CBE503B7
Public Function EliminaRHCuenta(psPersCod As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo EliminaRHCuentaErr

    'your code goes here...

    Exit Function
EliminaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "NAperturarCuentaRRHH:EliminaRHCuenta Method")
End Function

'Asigna una cuenta a un RRHH para que se le abone sus pagos
'##ModelId=3A96CBE503DF
Public Function ModificaRHCuenta(psPersCod As String, psCodTarj As String, psCodCta As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRHCuentaErr

    'your code goes here...

    Exit Function
ModificaRHCuentaErr:
    Call RaiseError(MyUnhandledError, "NAperturarCuentaRRHH:ModificaRHCuenta Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D20281
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
