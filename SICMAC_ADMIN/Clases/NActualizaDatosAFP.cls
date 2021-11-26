VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosAFP"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A94013202B3"
'Actualiza los Datos de la AFP
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902CF0178
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902CF0255
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902CF0237
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

'Agrega una nueva APF
'##ModelId=3A96A6320305
Public Function AgregaAFP(psRHAFPPersCod As String, psRHAFPAbreviatura As String, pnRHAFPPrima As Double, pnRHAFPVariable As Double, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAFPErr

    'your code goes here...

    Exit Function
AgregaAFPErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosAFP:AgregaAFP Method")
End Function

'Elimina  una AFP
'##ModelId=3A96A632032D
Public Function EliminaAFP(psRHAFPPersCod As String) As Boolean
    Dim oAFP As DActualizaDatosAFP
    Set oAFP = New DActualizaDatosAFP
    On Error GoTo EliminaAFPErr
    
    If oAFP.EliminaAFP(psRHAFPPersCod) Then
        EliminaAFP = True
    Else
        EliminaAFP = False
    End If
    
    Set oAFP = Nothing
    Exit Function
EliminaAFPErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosAFP:EliminaAFP Method")
End Function

'Actualiza AFP
'##ModelId=3A96A6320356
Public Function ModificaAFP(psRHAFPPersCod As String, psRHAFPAbreviatura As String, pnRHAFPPrima As Double, pnRHAFPVariable As Double, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAFPErr
    Dim oAFP As DActualizaDatosAFP
    Set oAFP = New DActualizaDatosAFP
 
    If oAFP.ModificaAFP(psRHAFPPersCod, psRHAFPAbreviatura, pnRHAFPPrima, pnRHAFPVariable, psUltimaActualizacion) Then
        ModificaAFP = True
    Else
        ModificaAFP = False
    End If
    
    Set oAFP = Nothing

    Exit Function
ModificaAFPErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosAFP:ModificaAFP Method")
End Function

Public Function ReporteAFP(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo ModificaAFPErr
    Dim lsCodigo As String * 15
    Dim lsDescripcion As String * 52
    Dim lsMontoVar As String * 10
    Dim lsMontoPri As String * 10
    Dim lsMovNro As String * 27
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsCadena As String
    Dim oAFP As DActualizaDatosAFP
    Set oAFP = New DActualizaDatosAFP
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnI As Integer
    
    lnPagina = 0
    lnItem = 0
    lsCadena = ""
    
    Set rs = oAFP.GetAFPs
    If Not (rs.EOF And rs.BOF) Then
    lsCadena = lsCadena & CabeceraPagina("Porcentaje de Descuento de AFP (%)", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
    lsCadena = lsCadena & Encabezado("Codigo;8; ;7;Descripcion;20; ;30;Prima;13;Variable;13;MovNro;10; ;20;", lnItem)
        While Not rs.EOF
            RSet lsCodigo = rs.Fields(1)
            lsDescripcion = rs.Fields(2)
            RSet lsMontoPri = rs.Fields(4)
            RSet lsMontoVar = rs.Fields(5)
            lsMovNro = rs.Fields(6)
            lsCadena = lsCadena & lsCodigo & "  " & lsDescripcion & "  " & lsMontoVar & "  " & lsMontoPri & "  " & lsMovNro & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Porcentaje de Descuento de AFP (%)", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8; ;3;Descripcion;20; ;30;Monto;13;MovNro;10; ;20;", lnItem)
            End If
            rs.MoveNext
        Wend
    End If
    
    Set oAFP = Nothing
    ReporteAFP = lsCadena
    
    Exit Function
ModificaAFPErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosAFP:ModificaAFP Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902CF018C
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
