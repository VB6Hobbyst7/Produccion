VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DProcesarPlanilla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9DAEF40061"
'Actualiza Planilla
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030A0213
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB9030A03D6
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB9030A039A
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
'##ModelId=3A9DB73C03A9
Public Function AgregaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRRHHPago As String, pdRRHHInicio As Date, pdRRHHFin As Date, psPlaInsDes As String) As Boolean
    On Error GoTo AgregaPlanillaErr

    'your code goes here...

    Exit Function
AgregaPlanillaErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:AgregaPlanilla Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB73C03E5
Public Function EliminaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String) As Boolean
    On Error GoTo EliminaPlanillaErr

    'your code goes here...

    Exit Function
EliminaPlanillaErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:EliminaPlanilla Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB73D002F
Public Function ModificaPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRRHHPago As String, pdRRHHInicio As Date, pdRRHHFin As Date, psPlaInsDes As String) As Boolean
    On Error GoTo ModificaPlanillaErr

    'your code goes here...

    Exit Function
ModificaPlanillaErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:ModificaPlanilla Method")
End Function

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9DB73D0061
Public Function AgregaPlanillaDetalle(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String, psSistPensiones As String, pnPorMesTrabajado As Double, psComentario As String) As Boolean
    On Error GoTo AgregaPlanillaDetalleErr

    'your code goes here...

    Exit Function
AgregaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:AgregaPlanillaDetalle Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB73D009D
Public Function EliminaPlanillaDetalle(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String) As Boolean
    On Error GoTo EliminaPlanillaDetalleErr

    'your code goes here...

    Exit Function
EliminaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:EliminaPlanillaDetalle Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB73D00CF
Public Function ModificaPlanillaDetalle(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo ModificaPlanillaDetalleErr

    'your code goes here...

    Exit Function
ModificaPlanillaDetalleErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:ModificaPlanillaDetalle Method")
End Function

'Agrega una nueva Proyeccion Laboral
'##ModelId=3A9DB73D010B
Public Function AgregaPlanillaDetalleConcepto(psRRHHPeriodo As String, psPlanillaTpoCod As String, psPersCod As String, psConceptoCod As String, pnValor As Double, pnMonto As Double) As Boolean
    On Error GoTo AgregaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
AgregaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:AgregaPlanillaDetalleConcepto Method")
End Function

'Elimina una proyeccion
'##ModelId=3A9DB73D013E
Public Function EliminaPlanillaDetalleConcepto(psRHPersCod As String, pdRHProyeccionFecha As Date) As Boolean
    On Error GoTo EliminaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
EliminaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:EliminaPlanillaDetalleConcepto Method")
End Function

'Actualiza una proyeccion
'##ModelId=3A9DB73D017A
Public Function ModificaPlanillaDetalleConcepto(psRHPersCod As String, pdRHProyeccionFecha As Date, psUltimaActualizacion As String, psRHProyeccionCargoCod As String) As Boolean
    On Error GoTo ModificaPlanillaDetalleConceptoErr

    'your code goes here...

    Exit Function
ModificaPlanillaDetalleConceptoErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:ModificaPlanillaDetalleConcepto Method")
End Function

'Devuelve las planillas de existentes
'##ModelId=3A9EA6C8027C
Public Function GetPlanillas(Optional psGrupo As String = "") As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim sqlPAux As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlPAux = ""
    If psGrupo <> "" Then sqlPAux = " Where cRHPlanillaCod like '" & psGrupo & "%'"
    
    sqlP = "Select cRHPlanillaCod Codigo , cRHPlanillaDescripcion Descrip from RHPlanillaTabla " & sqlPAux
    If oCon.AbreConexion Then
        Set GetPlanillas = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:GetPlanillas Method")
End Function

'Devuelve una planillaespecifica
'##ModelId=3A9EA6DC0112
Public Function GetPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String) As Recordset
    On Error GoTo GetPlanillaErr

    'your code goes here...

    Exit Function
GetPlanillaErr:
    Call RaiseError(MyUnhandledError, "DProcesarPlanilla:GetPlanilla Method")
End Function

#If DebugMode Then
    '##ModelId=3AB9030A026D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
