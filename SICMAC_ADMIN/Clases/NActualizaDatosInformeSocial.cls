VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosInformeSocial"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96B0390194"
'Actualiza Datos de Informe Social del RRHH de la Empresa
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D60363
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D7001B
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D70007
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
'##ModelId=3A96B0B0022B
Public Function AgregaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String, pnRHInfSocial As Integer, ptRHInfSocialTexto As String, psRHInfSocialComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaInformeSocialErr
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    
    oInf.AgregaInformeSocial psRHPersCod, pdRHInfSocialFecha, pnRHInfSocial, ptRHInfSocialTexto, psRHInfSocialComentario, psUltimaActualizacion
        
    Set oInf = Nothing
    Exit Function
AgregaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosInformeSocial:AgregaInformeSocial Method")
End Function

'Agrega un nuevo Informe Social a la Persona
'##ModelId=3A96B0B0022B
Public Function ReporteInformeSocial(psRHPersCod As String, pnRHInfSocial As Integer, ByVal psTitulo As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo AgregaInformeSocialErr
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    Dim rsI As ADODB.Recordset
    Dim lsCadena As String
    Set rsI = New ADODB.Recordset
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsFecha As String * 30
    Dim lsComentario As String * 80
    
    Set rsI = oInf.GetInformesSociales(psRHPersCod, pnRHInfSocial)
    
    If Not (rsI.EOF And rsI.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Fecha;15; ;10;Resumen;20; ;75;", lnItem)
        While Not rsI.EOF
            lsFecha = rsI.Fields(0)
            lsComentario = rsI.Fields(1)
            lnItem = lnItem + 1
            lsCadena = lsCadena & Space(5) & lsFecha & lsComentario & oImpresora.gPrnSaltoLinea
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Fecha;15; ;10;Resumen;20; ;75;", lnItem)
            End If
            rsI.MoveNext
        Wend
    End If
    Set oInf = Nothing
    
    ReporteInformeSocial = lsCadena
    Exit Function
AgregaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosInformeSocial:AgregaInformeSocial Method")
End Function

'Elimina  un Informe Social de la Persona
'##ModelId=3A96B0B002C2
Public Function EliminaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String) As Boolean
    On Error GoTo EliminaInformeSocialErr
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    
    oInf.EliminaInformeSocial psRHPersCod, pdRHInfSocialFecha
    
    Set oInf = Nothing
    Exit Function
EliminaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosInformeSocial:EliminaInformeSocial Method")
End Function

'Actualiza un Informe Social de la Persona
'##ModelId=3A96B0B002EA
Public Function ModificaInformeSocial(psRHPersCod As String, pdRHInfSocialFecha As String, pnRHInfSocial As Integer, ptRHInfSocialTexto As String, psRHInfSocialComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaInformeSocialErr
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    
    oInf.ModificaInformeSocial psRHPersCod, pdRHInfSocialFecha, pnRHInfSocial, ptRHInfSocialTexto, psRHInfSocialComentario, psUltimaActualizacion
        
    Set oInf = Nothing
    Exit Function
ModificaInformeSocialErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosInformeSocial:ModificaInformeSocial Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D60377
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
