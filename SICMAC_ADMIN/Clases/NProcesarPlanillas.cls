VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NProcesarPlanillas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9BEB120345"
'Actualiza Planilla
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902DE0397
    Private mlClassDebugID As Long
#End If

'##ModelId=3A9EA60C03BC
Public theNInterprete As NInterprete

'##ModelId=3AB902DF00E5
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902DF00C7
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

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9DB16C0327
Public Function AgregaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRRHHPago As String, pdRRHHInicio As Date, pdRRHHFin As Date, psPlaInsDes As String) As Boolean
    On Error GoTo AgregaPlanillaErr

    'your code goes here...

    Exit Function
AgregaPlanillaErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:AgregaPlanilla Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB16C034F
Public Function EliminaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String) As Boolean
    On Error GoTo EliminaPlanillaErr

    'your code goes here...

    Exit Function
EliminaPlanillaErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:EliminaPlanilla Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB16C0381
Public Function ModificaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRRHHPago As String, pdRRHHInicio As Date, pdRRHHFin As Date, psPlaInsDes As String) As Boolean
    On Error GoTo ModificaPlanillaErr

    'your code goes here...

    Exit Function
ModificaPlanillaErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:ModificaPlanilla Method")
End Function

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9DB19201E1
Public Function AgregaPlanillaDetalle(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String, psSistPensiones As String, pnPorMesTrabajado As Double, psComentario As String) As Boolean
    On Error GoTo AgregaPlanillaDetalleErr

    'your code goes here...

    Exit Function
AgregaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:AgregaPlanillaDetalle Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB1920209
Public Function EliminaPlanillaDetalle(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String) As Boolean
    On Error GoTo EliminaPlanillaDetalleErr

    'your code goes here...

    Exit Function
EliminaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:EliminaPlanillaDetalle Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB1920231
Public Function ModificaPlanillaDetalle(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo ModificaPlanillaDetalleErr

    'your code goes here...

    Exit Function
ModificaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:ModificaPlanillaDetalle Method")
End Function

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9DB1C2003C
Public Function AgregaPlanillaDetalleConcepto(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String, psConceptoCod As String, pnValor As Double, pnMonto As Double) As Boolean
    On Error GoTo AgregaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
AgregaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:AgregaPlanillaDetalleConcepto Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB1C20064
Public Function EliminaPlanillaDetalleConcepto(psRHPersCod As String, pdRHProyeccionFecha As Date) As Boolean
    On Error GoTo EliminaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
EliminaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:EliminaPlanillaDetalleConcepto Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB1C2008C
Public Function ModificaPlanillaDetalleConcepto(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo ModificaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
ModificaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:ModificaPlanillaDetalleConcepto Method")
End Function

'##ModelId=3A9DBB4F031D
Public Function ProcesarPlanilla() As Recordset
    On Error GoTo ProcesarPlanillaErr

    'your code goes here...

    Exit Function
ProcesarPlanillaErr:
    Call RaiseError(MyUnhandledError, "NProcesarPlanillas:ProcesarPlanilla Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902DE03BF
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
