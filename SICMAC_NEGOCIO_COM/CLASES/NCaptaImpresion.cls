VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCaptaImpresion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7A6853020C"
'Genera los registros de Firmas, las cartas de contrato y las constancias de
'plazo fijo, saca las plantillas de la base de datos de acuerdo a un codigo
'que se le envía
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3A835D5E0029
    Private mlClassDebugID As Long
#End If

'##ModelId=3A7B7F2C0285
Public theNImprimir As NImprimir

'##ModelId=3A7A69160086
'Public theNCapMantenimiento As NCapMantenimiento

'Obtiene la plantilla de la carta de Convenio - Tarjeta Magnetica, y le asigna
'valores a sus campos.
'##ModelId=3A7A68F203AF
''Public Function ImpCartaConvenioTarjeta(psCodCta As String, pClientesDocTpoRela As Recordset) As String
'    On Error GoTo ImpCartaConvenioTarjetaErr
'
'    'your code goes here...
'
'    Exit Function
'ImpCartaConvenioTarjetaErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpCartaConvenioTarjeta Method")
'End Function

''Obtiene la plantilla de la carta de Registro de Firmas y le asigna valores a
''sus campos.
''##ModelId=3A7A690202F4
'Public Function ImpRegFirmas(psCodCta As String, ClientesDocTpoRela As Recordset) As String
'    On Error GoTo ImpRegFirmasErr
'
'    'your code goes here...
'
'    Exit Function
'ImpRegFirmasErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpRegFirmas Method")
'End Function

'Obtiene la plantilla del extracto y le asigna valores a sus campos, y los
'movimientos.
'##ModelId=3A7B8032023C
'Public Function ImpExtracto(psCodCta As String, pClientesDocTpoRela As Recordset, pMovimientos As Recordset, psSaldosDisp As String, psSaldosCnt As String) As String
'    On Error GoTo ImpExtractoErr
'
'    'your code goes here...
'
'    Exit Function
'ImpExtractoErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpExtracto Method")
'End Function

'Obtiene la plantilla de la boleta y le asigna valores
'##ModelId=3A7B81300179
'Public Function ImpBoleta(psCodCta As String, pClientesDocTpoRela As Recordset, psSaldosDisp As String, psSaldosCnt As String, pnNroExtracto As String, psTtulo As String, psDescripcion As String, psLinea3 As String, psLinea4 As String, psUser As String, psUserRem As String) As String
'    On Error GoTo ImpBoletaErr
'
'    'your code goes here...
'
'    Exit Function
'ImpBoletaErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpBoleta Method")
'End Function

'Obtiene la plantilla de la boleta de extorno y le asigna valores
'##ModelId=3A8AF48502E7
'Public Function ImpBoletaExtorno(psCodCta As String, pClientesDocTpoRela As Recordset, psSaldosDisp As String, psSaldosCnt As String, pnNroExtracto As String, psTtulo As String, psDescripcion As String, psLinea3 As String, psLinea4 As String, psUser As String, psUserRem As String) As String
'    On Error GoTo ImpBoletaExtornoErr
'
'    'your code goes here...
'
'    Exit Function
'ImpBoletaExtornoErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpBoletaExtorno Method")
'End Function

'Obtiene la plantilla de las ordenes de pago asigna valores a sus campos
'##ModelId=3A7B819F024B
Public Function ImpChequera(psCuenta As String, psNombre As String, psDocumento As String) As String
    On Error GoTo ImpChequeraErr

    'your code goes here...

    Exit Function
ImpChequeraErr:
    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpChequera Method")
End Function

'Obtiene la plantilla de impresiones de simulacion agrega los datos
'##ModelId=3A7CF19A016F
'Public Function ImpSimulacion(psCadena As String, psCodRep As String, pMovimiento As Recordset) As String
'    On Error GoTo ImpSimulacionErr
'
'    'your code goes here...
'
'    Exit Function
'ImpSimulacionErr:
'    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpSimulacion Method")
'End Function

'##ModelId=3A818B4203B5
Public Function ImpDuplicadoPlazoFijo(psCodCta As String) As String
    On Error GoTo ImpDuplicadoPlazoFijoErr

    'your code goes here...

    Exit Function
ImpDuplicadoPlazoFijoErr:
    Call RaiseError(MyUnhandledError, "NCaptaImpresion:ImpDuplicadoPlazoFijo Method")
End Function

'##ModelId=3A835D5E021E
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3A835D5E028C
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3A835D5E00BF
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
