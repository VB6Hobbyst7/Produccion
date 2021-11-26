VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosAFP"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A969DB000A1"
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

'Elimina  una AFP
'##ModelId=3A969FF60198
Public Function EliminaAFP(psRHAFPPersCod As String) As Boolean
    Dim oCon As DConecta
    Dim sqlA As String
    Set oCon = New DConecta
    On Error GoTo EliminaAFPErr

    sqlA = "Delete " & lsServerAdministracion & "RHAFP where cRHAFPPersCod = '" & psRHAFPPersCod & "'"
    If oCon.AbreConexion() Then
        oCon.Ejecutar sqlA
        oCon.CierraConexion
        EliminaAFP = True
    Else
        EliminaAFP = False
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaAFPErr:
    EliminaAFP = False
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:EliminaAFP Method")
End Function

'Actualiza AFP
'##ModelId=3A969FF601CA
Public Function ModificaAFP(psRHAFPPersCod As String, psRHAFPAbreviatura As String, pnRHAFPPrima As Double, pnRHAFPVariable As Double, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo ModificaAFPErr
    
    sqlC = " Update " & lsServerAdministracion & "RHAFP" _
         & " Set cRHAFPAbreviatura = '" & psRHAFPAbreviatura & "', nRHAFPPrima = " & pnRHAFPPrima & ", nRHAFPVariable = " & pnRHAFPVariable & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " where cRHAFPPersCod = '" & psRHAFPPersCod & "'"
    
    If oCon.AbreConexion() Then
        oCon.Ejecutar sqlC
        
        ModificaAFP = True
        oCon.CierraConexion
    Else
        ModificaAFP = False
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaAFPErr:
    ModificaAFP = False
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:ModificaAFP Method")
End Function

'Devuelve las AFP Activas
'##ModelId=3A96A0130009
Public Function GetAFP(Optional pbTipoArbol As Boolean = False) As Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlA As String
    On Error GoTo GetAFPErr
    
    If Not pbTipoArbol Then
        sqlA = " Select PE.cPersNombre, AF.cRHAFPPersCod  From " & lsServerAdministracion & "RHAFP AF" _
             & " Inner Join " & lsServerPersona & "Persona  PE On AF.cRHAFPPersCod = PE.cPersCod"
    Else
        sqlA = " Select cRHAFPAbreviatura Abre, PE.cPersNombre + ' - ' + AF.cRHAFPPersCod, Len(AF.cRHAFPPersCod) Nivel From " & lsServerAdministracion & "RHAFP AF" _
             & " Inner Join " & lsServerPersona & "Persona  PE On AF.cRHAFPPersCod = PE.cPersCod"
    End If
    
    If oCon.AbreConexion Then
        Set GetAFP = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAFPErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetAFP Method")
End Function

'Devuelve las AFP Activas
'##ModelId=3A96A0130009
Public Function GetAFPs() As Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlA As String
    On Error GoTo GetAFPsErr
    
    If oCon.AbreConexion() Then
        sqlA = " Select '', cRHAFPPersCod as Codigo, cPersNombre as Nombre, cRHAFPAbreviatura as Abre," _
             & " Case nRHAFPPrima when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHAFPPrima)) end as Prima, Case nRHAFPVariable when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHAFPVariable)) end as ComVar, RA.cUltimaActualizacion as Actualizacion from " & lsServerAdministracion & "RHAFP RA" _
             & " Inner Join " & lsServerPersona & "Persona PE On RA.cRHAFPPersCod = PE.cPersCod" _
             & " Order by PE.cPersNombre"
        Set GetAFPs = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetAFPsErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetAFP Method")
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetValorAFP(psRHAFPPersCod As String, pbRHAFPPrima As Boolean) As Double
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
        
        GetValorAFP = rsA!Monto
    Else
        GetValorAFP = -1
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
