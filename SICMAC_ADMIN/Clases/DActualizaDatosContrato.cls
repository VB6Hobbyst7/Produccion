VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosContrato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96BDF50122"
'Actualiza en Base los datos del Contrato del RRHH
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902F70053
    Private mlClassDebugID As Long
#End If

'Elimina un contrato de un RRHH
'##ModelId=3A96F7B3031B
Public Function EliminaContrato(psRHPersCod As String, psRHContratoNro As String, Optional pdRHContratoFecha As Date) As Boolean
    On Error GoTo EliminaContratoErr

    'your code goes here...

    Exit Function
EliminaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:EliminaContrato Method")
End Function

'##ModelId=3AB902F701A8
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F7016C
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
'##ModelId=3A96F7B302DF
'Public Function AgregaContrato(psRHPersCod As String, psRHContratoNro As String, ptRHContratoTexto As String, psRHContratoComentario As String, psUltimaActualizacion As String, pdSistema As Date, pdRHContratoInicio As String, pdRHContratoFin As String, pnRHSueldoMonto As Currency) As Boolean
Public Function AgregaContrato(psRHPersCod As String, psRHContratoNro As String, ptRHContratoTexto As String, psUltimaActualizacion As String, pnRHContratoTpo As Long) As Boolean
    On Error GoTo AgregaContratoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Insert RHContrato (cPersCod,cRHContratoNro,tRHContratoTexto,cUltimaActualizacion,nRHContratoTpo)" _
         & " values('" & psRHPersCod & "','" & psRHContratoNro & "','" & ptRHContratoTexto & "','" & psUltimaActualizacion & "'," & pnRHContratoTpo & ")"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96F7B302DF
Public Function AgregaContratoDet(psRHPersCod As String, psRHContratoNro As String, pdSistema As String, psRHContratoInicio As String, psRHContratoFin As String, psRHContratoComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaContratoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
   
    sqlC = " Insert RHContratoDet (cPersCod,cRHContratoNro,dRHContratoFecha,dRHContratoInicio,dRHContratoFin,cRHContratoComentario,cUltimaActualizacion)" _
         & " values('" & psRHPersCod & "','" & psRHContratoNro & "','" & pdSistema & "','" & psRHContratoInicio & "','" & psRHContratoFin & "','" & psRHContratoComentario & "','" & psUltimaActualizacion & "')"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96F7B302DF
Public Function AgregaSueldo(psPersCod As String, psRHSueldoFecha As String, pnRHSueldoMonto As Currency, psRHSueldoComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaContratoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Insert RHSueldo (cPersCod,dRHSueldoFecha,nRHSueldoMonto,cRHSueldoComentario,cUltimaActualizacion)" _
         & " values('" & psPersCod & "','" & psRHSueldoFecha & "'," & pnRHSueldoMonto & ",'" & psRHSueldoComentario & "','" & psUltimaActualizacion & "')"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96F7B302DF
Public Function GetContratos(psRHPersCod As String) As ADODB.Recordset
    On Error GoTo AgregaContratoErr
    
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta

    sqlC = "Select distinct RHC.cRHContratoNro, RHCD.dRHContratoComentario Con From RHContrato RHC Inner Join RHContratoDet RHCD On RHC.cPersCod = RHCD.cPersCod And RHC.cPersCod = RHCD.cPersCod And RHC.cRHContratoNro= RHCD.cRHContratoNro where RHC.cPersCod = '" & psRHPersCod & "' order by RHC.cRHContratoNro Desc"
    
    If oCon.AbreConexion Then
        Set GetContratos = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96F7B302DF
Public Function GetContrato(psRHPersCod As String, psContratoCod As String) As ADODB.Recordset
    On Error GoTo AgregaContratoErr
    
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select Top 1  tRHContratoTexto Texto, dRHContratoFecha Fecha, dRHContratoInicio dIni,dRHContratoFin dFin, dRHContratoComentario Comen from RHContrato RHC " _
         & " Inner Join RHContratoDet RHCD On RHCD.cPersCod = RHC.cPersCod ANd RHCD.cRHContratoNro = RHC.cRHContratoNro" _
         & " where RHC.cPerscod = '" & psRHPersCod & "' And RHC.cRHContratoNro = '" & psContratoCod & "' Order By dRHContratoFecha Desc"
    If oCon.AbreConexion Then
        Set GetContrato = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function

'Agrega un nuevo contrato de un RRHH
'##ModelId=3A96F7B302DF
Public Function GetSueldoContrato(psRHPersCod As String) As ADODB.Recordset
    On Error GoTo AgregaContratoErr
    
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select nRHSueldoMonto Sueldo, cUltimaActualizacion Mov from rhsueldo" _
         & " Where cPerscod = '" & psRHPersCod & "' Order by dRHSueldoFecha Desc"
    If oCon.AbreConexion Then
        Set GetSueldoContrato = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:AgregaContrato Method")
End Function


'Actualiza un contrato de un RRHH
'##ModelId=3A96F7B30357
Public Function ModificaContrato(psRHPersCod As String, psRHContratoNro As String, ptRHContratoTexto As String, psRHContratoComentario As String, psUltimaActualizacion As String, pdSistema As Date, pdRHContratoInicio As Date, pdRHContratoFin As Date) As Boolean
    On Error GoTo ModificaContratoErr

    'your code goes here...

    Exit Function
ModificaContratoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:ModificaContrato Method")
End Function

'Modifica el Sueldo de un RRHH
'##ModelId=3A96F7B30393
Public Function ModificaSueldo(psRHPersCod As String, pdRHSueldoFecha As Date, pnRHSueldoMonto As Currency, psRHSueldoComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaSueldoErr
    
    
    Exit Function
ModificaSueldoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:ModificaSueldo Method")
End Function


'Modifica el Sueldo de un RRHH
'##ModelId=3A96F7B30393
Public Function GeCodContrato(psRHPersCod As String) As String
    On Error GoTo ModificaSueldoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlC = "Select Top 1 cRHContratoNro from rhcontrato where cPersCod = '" & psRHPersCod & "' Order by cRHContratoNro Desc "
    
    If oCon.AbreConexion Then
        Set rs = oCon.Ejecutar(sqlC)
        If rs.EOF And rs.BOF Then
            GeCodContrato = "000001"
        Else
            GeCodContrato = Format(CCur(rs!cRHContratoNro) + 1, "000000")
        End If
        oCon.CierraConexion
    End If
    
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
ModificaSueldoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosContrato:ModificaSueldo Method")
End Function


#If DebugMode Then
    '##ModelId=3AB902F700AE
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
