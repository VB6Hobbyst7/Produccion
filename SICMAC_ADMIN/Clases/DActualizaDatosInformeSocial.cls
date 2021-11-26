VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosInformeSocial"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96B08C03CE"
'Actualiza en Base Datos de Informe Social del RRHH de la Empresa
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902F5015F
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902F502A9
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F50277
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

'Agrega un nuevo Informe Social a la Persona
'##ModelId=3A96B56600B3
Public Function AgregaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String, pnRHInfSocial As Integer, ptRHInfSocialTexto As String, psRHInfSocialComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Insert RHInformeSocial (cPersCod,dRHInfSocialFecha,nRHInfSocialTpo,tRHInfSocialTexto,cRHInfSocialComentario,cUltimaActualizacion)" _
         & " Values ('" & psRHPersCod & "','" & pdRHInfSocialFecha & "'," & Trim(Str(pnRHInfSocial)) & ",'" & ptRHInfSocialTexto & "','" & psRHInfSocialComentario & "','" & psUltimaActualizacion & "')"

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:AgregaInformeSocial Method")
End Function

'Elimina  un Informe Social de la Persona
'##ModelId=3A96B56600EF
Public Function EliminaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String) As Boolean
    On Error GoTo EliminaInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Delete RHInformeSocial " _
         & " Where cPersCod = '" & psRHPersCod & "' And dRHInfSocialFecha = '" & pdRHInfSocialFecha & "'"

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:EliminaInformeSocial Method")
End Function

'Actualiza un Informe Social de la Persona
'##ModelId=3A96B566012B
Public Function ModificaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String, pnRHInfSocial As Integer, ptRHInfSocialTexto As String, psRHInfSocialComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Update RHInformeSocial" _
         & " Set nRHInfSocialTpo = " & Trim(Str(pnRHInfSocial)) & ", tRHInfSocialTexto = '" & ptRHInfSocialTexto & "', cRHInfSocialComentario = '" & psRHInfSocialComentario & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cPersCod = '" & psRHPersCod & "' And dRHInfSocialFecha = '" & pdRHInfSocialFecha & "'"

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:ModificaInformeSocial Method")
End Function

'Devuelve los Informes sociales de un RRHH
'##ModelId=3A96B57D0160
Public Function GetInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String, pnTipoInf As Integer) As Recordset
    On Error GoTo GetInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select tRHInfSocialTexto Texto, cRHInfSocialComentario Comentario, nRHInfSocialTpo Tipo from rhInformesocial" _
         & " Where cPersCod = '" & psRHPersCod & "' And dRHInfSocialFecha = '" & pdRHInfSocialFecha & "'"

    If oCon.AbreConexion Then
        Set GetInformeSocial = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:GetInformeSocial Method")
End Function


'Devuelve los Informes sociales de un RRHH
'##ModelId=3A96B57D0160
Public Function GetInformesSociales(psRHPersCod As String, pnTipoInf As Integer) As Recordset
    On Error GoTo GetInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select Convert(varchar(10),dRHInfSocialFecha,103) + ' ' + Convert(varchar(10),dRHInfSocialFecha,108) Fecha, cRHInfSocialComentario Comentario, 1 Nivel " _
         & " FROM rhInformesocial Where cPersCod = '" & psRHPersCod & "'"

    If oCon.AbreConexion Then
        Set GetInformesSociales = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:GetInformeSocial Method")
End Function

'Devuelve un informe social de un RRHH
'##ModelId=3A96B57D019C
Public Function GetValorInformeSocial(psRHPersCod As String, psFecha As String) As ADODB.Recordset
    On Error GoTo GetValorInformeSocialErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select nRHInfSocialTpo Tpo, tRHInfSocialTexto Texto , cRHInfSocialComentario Coment from rhInformesocial " _
         & " Where cPersCod = '" & psRHPersCod & "' And  dRHInfSocialFecha = '" & psFecha & "'"

    If oCon.AbreConexion Then
        Set GetValorInformeSocial = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetValorInformeSocialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosInformeSocial:GetValorInformeSocial Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902F501B9
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
