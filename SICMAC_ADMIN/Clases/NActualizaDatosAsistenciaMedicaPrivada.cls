VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaAsistMedicaPriv"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96A1C700F6"
'Actualiza los Datos de la Asistencia Medica Privada
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D502A3
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D6001A
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D60006
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

'Agrega un nuevo Item de la Asistencia Medica Privada
'##ModelId=3A96A61703CF
'Public Function AgregaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
'    Dim oAsistencia As DActualizaAsistMedicaPrivada
'    Set oAsistencia = New DActualizaAsistMedicaPrivada
'    On Error GoTo AgregaAsisMedPrivErr
'
'    If oAsistencia.AgregaAsisMedPriv(psRHAsistMedPrivCod, psRHAsistMedPrivDescripcion, pnRHAsistMedPrivMonto, pnRHAsistMedPrivMontoDesc, psUltimaActualizacion) Then
'        AgregaAsisMedPriv = True
'    Else
'        AgregaAsisMedPriv = False
'    End If
'
'    Set oAsistencia = Nothing
'    Exit Function
'AgregaAsisMedPrivErr:
'    AgregaAsisMedPriv = False
'    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:AgregaAsisMedPriv Method")
'End Function

'Elimina  un Item de la Asistencia Medica Privada
'##ModelId=3A96A6180019
Public Function EliminaAsisMedPriv(psRHAsistMedPrivCod As String) As Boolean
    Dim oAsistencia As DActualizaAsistMedicaPrivada
    Set oAsistencia = New DActualizaAsistMedicaPrivada
    
    On Error GoTo EliminaAsisMedPrivErr

    If oAsistencia.EliminaAsisMedPriv(psRHAsistMedPrivCod) Then
        EliminaAsisMedPriv = True
    Else
        EliminaAsisMedPriv = False
    End If

    Set oAsistencia = Nothing
    Exit Function
EliminaAsisMedPrivErr:
    EliminaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:EliminaAsisMedPriv Method")
End Function

'Elimina  un Item de la Asistencia Medica Privada
'##ModelId=3A96A6180019
Public Function TipoUsado(psRHAsistMedPrivCod As String) As Boolean
    Dim oAsistencia As DActualizaAsistMedicaPrivada
    Set oAsistencia = New DActualizaAsistMedicaPrivada
    
    On Error GoTo EliminaAsisMedPrivErr
    
    TipoUsado = oAsistencia.TipoUsado(psRHAsistMedPrivCod)
    
    Set oAsistencia = Nothing
    Exit Function
EliminaAsisMedPrivErr:
    TipoUsado = False
    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:EliminaAsisMedPriv Method")
End Function


'Actualiza un Item de la Asistencia Medica Privada
'##ModelId=3A96A618004B
Public Function ModificaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
    Dim oAsistencia As DActualizaAsistMedicaPrivada
    Set oAsistencia = New DActualizaAsistMedicaPrivada
    
    On Error GoTo ModificaAsisMedPrivErr

    If oAsistencia.ModificaAsisMedPriv(psRHAsistMedPrivCod, psRHAsistMedPrivDescripcion, pnRHAsistMedPrivMonto, pnRHAsistMedPrivMontoDesc, psUltimaActualizacion) Then
        ModificaAsisMedPriv = True
    Else
        ModificaAsisMedPriv = False
    End If

    Exit Function
ModificaAsisMedPrivErr:
    ModificaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function


''''Public Function ModificaAsisMedPriv(psnCodCatalogo As String, pscCodCatEPS As String, psRHAsistMedPrivDescripcion As String, pnnCantPer As Integer, nCostoTotal As Currency, pnCostoIGV As Currency, pnCostoNeto As Currency, psUltimaActualizacion As String, pnAdeQuin As Currency) As Boolean
''''    Dim oAsistencia As DActualizaAsistMedicaPrivada
''''    Set oAsistencia = New DActualizaAsistMedicaPrivada
''''
''''    On Error GoTo ModificaAsisMedPrivErr
''''
''''    If oAsistencia.ModificaAsisMedPriv(psnCodCatalogo, pscCodCatEPS, psRHAsistMedPrivDescripcion, pnnCantPer, nCostoTotal, pnCostoIGV, pnCostoNeto, psUltimaActualizacion, pnAdeQuin) Then
''''        ModificaAsisMedPriv = True
''''    Else
''''        ModificaAsisMedPriv = False
''''    End If
''''
''''    Exit Function
''''ModificaAsisMedPrivErr:
''''    ModificaAsisMedPriv = False
''''    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
''''End Function



'Actualiza un Item de la Asistencia Medica Privada
'##ModelId=3A96A618004B
Public Function GetReporte(pgsNomAge As String, pgsEmpresa As String, pdFecSis As Date) As String
    Dim oAsistencia As DActualizaAsistMedicaPrivada
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oAsistencia = New DActualizaAsistMedicaPrivada
    Set rs = oAsistencia.GetAsisMedPriv
    Dim lsCodigo As String * 8
    Dim lsDescripcion As String * 52
    Dim lsMonto As String * 10
    Dim lsMontoDesc As String * 10
    Dim lsMovNro As String * 27
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    
    On Error GoTo ModificaAsisMedPrivErr
    
    lnPagina = 0
    lnItem = 0
    lsCadena = ""
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Asitencia Medica Privada", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8; ;3;Descripcion;20; ;30;Monto;13;Descuent;15;MovNro;15; ;15;", lnItem)
    
        While Not rs.EOF
            RSet lsCodigo = rs.Fields(1)
            lsDescripcion = rs.Fields(2)
            RSet lsMonto = rs.Fields(3)
            RSet lsMontoDesc = rs.Fields(4)
            lsMovNro = rs.Fields(5)
            lsCadena = lsCadena & lsCodigo & "  " & lsDescripcion & "  " & lsMonto & "  " & "  " & lsMontoDesc & "  " & lsMovNro & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Asitencia Medica Privada", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8; ;3;Descripcion;20; ;30;Monto;13;Descuent;15;MovNro;15; ;15;", lnItem)
            End If
            rs.MoveNext
        Wend
    End If
    
    GetReporte = lsCadena
    
    Exit Function
ModificaAsisMedPrivErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D50380
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function AgregaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
    Dim oAsistencia As DActualizaAsistMedicaPrivada
    Set oAsistencia = New DActualizaAsistMedicaPrivada
    On Error GoTo AgregaAsisMedPrivErr
    
    If oAsistencia.AgregaAsisMedPriv(psRHAsistMedPrivCod, psRHAsistMedPrivDescripcion, pnRHAsistMedPrivMonto, pnRHAsistMedPrivMontoDesc, psUltimaActualizacion) Then
        AgregaAsisMedPriv = True
    Else
        AgregaAsisMedPriv = False
    End If
    
    Set oAsistencia = Nothing
    Exit Function
AgregaAsisMedPrivErr:
    AgregaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "NActualizaDatosAsistenciaMedicaPrivada:AgregaAsisMedPriv Method")
End Function

