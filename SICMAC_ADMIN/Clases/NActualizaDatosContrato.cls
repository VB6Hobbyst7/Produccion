VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosContrato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96BDD9008B"
'Actualiza los datos del Contrato del RRHH
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D7030A
    Private mlClassDebugID As Long
#End If

'Elimina un contrato de un RRHH
'##ModelId=3A96BE1C03D1
Public Function EliminaContrato(psRHPersCod As String, psRHContratoNro As String, Optional pdRHContratoFecha As Date) As Boolean
    On Error GoTo EliminaContratoErr

    'your code goes here...

    Exit Function
EliminaContratoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosContrato:EliminaContrato Method")
End Function

'##ModelId=3AB902D703C9
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D703B5
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

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96BE1C03A9
Public Function AgregaContrato(psRHPersCod As String, psRHContratoNro As String, ptRHContratoTexto As String, psRHContratoComentario As String, psUltimaActualizacion As String, psSistema As String, psRHContratoInicio As String, psRHContratoFin As String, pnRHSueldoMonto As Currency, pnRHContratoTpo As Long) As Boolean
    On Error GoTo AgregaContratoErr
    Dim oCon As DActualizaDatosContrato
    Set oCon = New DActualizaDatosContrato
    Dim lsCodContrato As String
    
    If Not IsNumeric(psRHContratoNro) Then
        lsCodContrato = oCon.GeCodContrato(psRHPersCod)
        oCon.AgregaContrato psRHPersCod, lsCodContrato, ptRHContratoTexto, psUltimaActualizacion, pnRHContratoTpo
        oCon.AgregaContratoDet psRHPersCod, lsCodContrato, psSistema, psRHContratoInicio, psRHContratoFin, psRHContratoComentario, psUltimaActualizacion
    Else
        oCon.AgregaContratoDet psRHPersCod, psRHContratoNro, psSistema, psRHContratoInicio, psRHContratoFin, psRHContratoComentario, psUltimaActualizacion
    End If
    If pnRHSueldoMonto <> -1 Then oCon.AgregaSueldo psRHPersCod, psSistema, pnRHSueldoMonto, psRHContratoComentario, psUltimaActualizacion
     
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosContrato:AgregaContrato Method")
End Function

'Actualiza un contrato de un RRHH
'##ModelId=3A96BE1D0011
Public Function ModificaContrato(psRHPersCod As String, psRHContratoNro As String, ptRHContratoTexto As String, psRHContratoComentario As String, psUltimaActualizacion As String, pdSistema As Date, pdRHContratoInicio As Date, pdRHContratoFin As Date) As Boolean
    On Error GoTo ModificaContratoErr

    'your code goes here...

    Exit Function
ModificaContratoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosContrato:ModificaContrato Method")
End Function

'Modifica el Sueldo de un RRHH
'##ModelId=3A96C10403C7
Public Function ModificaSueldo(psRHPersCod As String, pdRHSueldoFecha As Date, pnRHSueldoMonto As Currency, psRHSueldoComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaSueldoErr

    'your code goes here...

    Exit Function
ModificaSueldoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosContrato:ModificaSueldo Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D70328
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
