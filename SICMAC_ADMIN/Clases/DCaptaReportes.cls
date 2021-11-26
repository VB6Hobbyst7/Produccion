VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCaptaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en Base la AFP
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'##ModelId=3AB902F20151
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F20114
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

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'Agrega una nueva APF
'##ModelId=3A969FF6015C
Public Function AgregaAFP(psRHAFPPersCod As String, psRHAFPAbreviatura As String, pnRHAFPPrima As Double, pnRHAFPVariable As Double, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAFPErr

    'your code goes here...

    Exit Function
AgregaAFPErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:AgregaAFP Method")
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetCaptaReportes(pbRHAFPPrima As Boolean, psRHAFPPersCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlA As String
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    
    On Error GoTo GetValorAFPErr

    If oCon.AbreConexion() Then
        If pbRHAFPPrima Then
            sqlA = " Select nRHAFPPrima as Monto From " & lsServerAdministracion & "RHAFP where cRHAFPPersCod  = '" & psRHAFPPersCod & "'"
        Else
            sqlA = " Select nRHAFPVariable as Monto From " & lsServerAdministracion & "RHAFP where cRHAFPPersCod  = '" & psRHAFPPersCod & "'"
        End If
        
        Set rsA = oCon.CargaRecordSet(sqlA)
        Set GetCaptaReportes = rsA
    End If
    
    rsA.Close
    Set rsA = Nothing
    Set oCon = Nothing
    Exit Function
GetValorAFPErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902F20056
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

