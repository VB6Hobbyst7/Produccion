VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosCargo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A94012F0219"
'Actualiza los cargos y niveles
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902CF0042
    Private mlClassDebugID As Long
#End If

'##ModelId=3A9454E90018
Public theNCaptaImpresion As NCaptaImpresion

'##ModelId=3AB902CF00EC
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902CF00D8
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

'Agrega un nuevo cargo o nivel
'##ModelId=3A942BD30337
Public Function AgregaCargo(psRHCargoCod As String, psRHCargoDescripcion As String, pnRHCargoSueldoMinimo As Double, pnRHCargoSueldoMaximo As Double, psRHCargoCategoria As String, psUltimaActualizacion As String, pnRHCargoOrden As Integer, pnGrado As Integer, pnCtrAsist, pnRHGrupoPlanea As Integer, pnRHGastoDirInd As Integer) As Boolean
    Dim oCargo As DActualizadatosCargo
    Set oCargo = New DActualizadatosCargo
    On Error GoTo AgregaCargoErr
    
    oCargo.AgregaCargo psRHCargoCod, psRHCargoDescripcion, pnRHCargoSueldoMinimo, pnRHCargoSueldoMaximo, psRHCargoCategoria, psUltimaActualizacion, pnRHCargoOrden, pnGrado, pnCtrAsist, pnRHGrupoPlanea, pnRHGastoDirInd
    
    Set oCargo = Nothing
    Exit Function
AgregaCargoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:AgregaCargo Method")
End Function

'Elimina  un nuevo cargo o nivel
'##ModelId=3A942BD3035F
Public Function EliminaCargo(psRHCargoCod As String) As Boolean
    Dim oCargo As DActualizadatosCargo
    Set oCargo = New DActualizadatosCargo
    On Error GoTo EliminaCargoErr

    oCargo.EliminaCargo psRHCargoCod
    
    Set oCargo = Nothing
    Exit Function
EliminaCargoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:EliminaCargo Method")
End Function

'Elimina  un nuevo cargo o nivel
'##ModelId=3A942BD3035F
Public Function VerificaCargoUsado(psRHCargoCod As String) As Boolean
    Dim oCargo As DActualizadatosCargo
    Set oCargo = New DActualizadatosCargo
    On Error GoTo EliminaCargoErr
    
    VerificaCargoUsado = oCargo.VerificaUsoCargo(psRHCargoCod)
    
    Set oCargo = Nothing
    Exit Function
EliminaCargoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:EliminaCargo Method")
End Function

'Elimina  un nuevo cargo o nivel
'##ModelId=3A942BD3035F
Public Function GetReporte(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteErr
    Dim oCargo As DActualizadatosCargo
    Dim rsC As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim oPrevio As Previo.clsPrevio
    Set oCargo = New DActualizadatosCargo
    Set rsC = New ADODB.Recordset
    Set rsC = oCargo.GetCargos(False, , True)
    Set oCargo = Nothing
    Set oPrevio = New Previo.clsPrevio
    
    
    Dim lsCodigo As String * 8
    Dim lsDescripcion As String * 50
    Dim SueMin As String * 15
    Dim SueMax As String * 15
    Dim Categoria As String * 12
    
    lsCadena = ""
    
    If Not (rsC.EOF And rsC.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Cargos y Niveles de la CMACT", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;22;Sueldo;15;Sueldo Max;15;Categoria;15; ;23;", lnItem)
              
        While Not rsC.EOF
            lsCodigo = rsC.Fields(1)
            lsDescripcion = rsC.Fields(2)
            SueMin = rsC.Fields(3)
            SueMax = IIf(IsNull(rsC.Fields(4)), "", rsC.Fields(4))
            Categoria = IIf(IsNull(rsC.Fields(6)), "", rsC.Fields(6))
            lsCadena = lsCadena & lsCodigo & lsDescripcion & SueMin & SueMax & Categoria & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Cargos y Niveles de la CMACT", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;22;Sueldo;15;Sueldo Max;15;Categoria;15; ;23;", lnItem)
           End If
           rsC.MoveNext
        Wend
    End If
        
    rsC.Close
    Set rsC = Nothing
    
    GetReporte = lsCadena
    
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:EliminaCargo Method")
End Function
    
'Actualiza un nuevo cargo o nivel
'##ModelId=3A942BD30387
Public Function ModificaCargo(psRHCargoCod As String, psRHCargoDescripcion As String, pnRHCargoSueldoMinimo As Integer, pnRHCargoSueldoMaximo As Integer, psRHCargoCategoria As String, psUltimaActualizacion As String, pnRHCargoOrden As Integer, pnGrado As Integer, pnCtrAsist As Integer, pnRHGrupoPlanea As Integer, pnRHGastoDirInd As Integer) As Boolean
    Dim oCargos As DActualizadatosCargo
    Set oCargos = New DActualizadatosCargo
    On Error GoTo ModificaCargoErr
    
    oCargos.ModificaCargo psRHCargoCod, psRHCargoDescripcion, pnRHCargoSueldoMinimo, pnRHCargoSueldoMaximo, psRHCargoCategoria, psUltimaActualizacion, pnRHCargoOrden, pnGrado, pnCtrAsist, pnRHGrupoPlanea, pnRHGastoDirInd
    
    Set oCargos = Nothing
    Exit Function
ModificaCargoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:ModificaCargo Method")
End Function

'Actualiza un nuevo cargo o nivel
'##ModelId=3A942BD30387
Public Function SetAreasCargo(prRS As ADODB.Recordset, psRHCargoCod As String) As Boolean
    Dim oCargos As DActualizadatosCargo
    Set oCargos = New DActualizadatosCargo
    On Error GoTo ModificaCargoErr
    
    SetAreasCargo = oCargos.SetAreasCargo(prRS, psRHCargoCod)
    
    Set oCargos = Nothing
    Exit Function
ModificaCargoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCargo:ModificaCargo Method")
End Function


#If DebugMode Then
    '##ModelId=3AB902CF0056
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
