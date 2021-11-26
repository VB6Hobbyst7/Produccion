VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHPrestamosAdm"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza Conceptos
Option Base 0
Option Explicit
Dim oCon As DConecta

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FC0097
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902FC01F5
Private Sub Class_Terminate()
    oCon.CierraConexion
    
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902FC01AF
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Set oCon = New DConecta
    oCon.AbreConexion
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub


'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function GetCreditosAdm(psPlaCod As String, psPeriodo As String) As ADODB.Recordset
    On Error GoTo GetFormulaErr
    Dim sqlD As String

    sqlD = " Select P.cPersNombre cNomPers , E.cPersCod, E.cRHCod cEmpCod, C.cCtaCod cCodCta, C.nMonto, C.cNroCuota, C.dFecVenc " _
         & " FROM Persona P" _
         & " INNER JOIN RHCreditosAdm C ON P.cPersCod = C.cPersCod" _
         & " INNER JOIN RRHH E ON E.cPersCod = C.cPersCod" _
         & " WHERE C.cPlaCod = '" & psPlaCod & "'" _
         & " And cPeriodo like '" & psPeriodo & "' And C.cEstado = '0' Order by P.cNomPers, C.cCodCta"
    
    Set GetCreditosAdm = oCon.CargaRecordSet(sqlD)
    
    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function


'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function GetCreditosAdmOtros(psPlaCod As String, psPeriodo As String) As ADODB.Recordset
    On Error GoTo GetFormulaErr
    Dim sqlD As String

    sqlD = " Select E.cPersCod, P.cPersNombre cNomPers, CO.cConsDescripcion + Space(75) + Rtrim(Ltrim(Str(C.nMoneda))) cCodCta, (Select P.cPersNombre + Space(75) + C.cPersCodInst From Persona P Where P.cPersCod = C.cPersCodInst)  Inst, Case When C.nTipCambio Is Null Then C.nMonto Else Round(C.nTipCambio * C.nMonto,2) End,  Case When C.nTipCambio Is Not Null Then C.nMonto Else 0 End, C.nTipCambio" _
         & " FROM Persona P" _
         & " INNER JOIN RHCreditosAdmOtros C ON P.cPersCod = C.cPersCod" _
         & " INNER JOIN RRHH E ON E.cPersCod = C.cPersCod" _
         & " INNER JOIN Constante CO ON CO.nConsValor = C.nMoneda And CO.nConsCod = " & gMoneda & "" _
         & " WHERE C.cPlaCod = '" & psPlaCod & "'" _
         & " And cPeriodo like '" & psPeriodo & "%' Order by P.cNomPers, C.nMoneda"
    
    Set GetCreditosAdmOtros = oCon.CargaRecordSet(sqlD)
    
    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function


'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function EliminaCreditosAdm(psPlaCod As String, psPeriodo As String) As Boolean
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    
    sqlD = " Delete RHCreditosAdm " _
         & " Where cPlaCod = '" & psPlaCod & "' And cPeriodo = '" & psPeriodo & "'"
    oCon.Ejecutar sqlD
    
    sqlD = " Delete RHConcepto " _
         & " Where cRHPlanillaCod = '" & psPlaCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdmin & "'"
    oCon.Ejecutar sqlD
    
    EliminaCreditosAdm = True
    Exit Function
GetFormulaErr:
    EliminaCreditosAdm = False
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function InsertaCreditosAdm(pcPersCod As String, pcPlaCod As String, pcPeriodo As String, pcCtaCod As String, pcNroCuota As String, pdFecVenc As Date, pnMonto As String, pnTipCambio As String, pcEstado As String, pcMovNro As String, pnMontoSoles As String) As Boolean
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlD = " Insert RHCreditosAdm (cPersCod,cPlaCod,cPeriodo,cCtaCod,cNroCuota,dFecVenc,nMonto,nTipCambio,cEstado,cMovNro)" _
         & " Values('" & pcPersCod & "','" & pcPlaCod & "','" & pcPeriodo & "','" & pcCtaCod & "','" & pcNroCuota & "','" & Format(pdFecVenc, gsFormatoFecha) & "'," & pnMonto & "," & pnTipCambio & ",'" & pcEstado & "','" & pcMovNro & "')"
    oCon.Ejecutar sqlD
    
    sqlD = " Select nRHConceptoValor From RHConcepto" _
         & " Where cRHPlanillaCod = '" & pcPlaCod & "' And cPersCod = '" & pcPersCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdmin & "'"
    Set rs = oCon.CargaRecordSet(sqlD)
    
    If rs.BOF And rs.EOF Then
        sqlD = " Insert RHConcepto (cRHPlanillaCod,cPersCod,cRHConceptoCod,nRHConceptoValor,cUltimaActualizacion)" _
             & " Values('" & pcPlaCod & "','" & pcPersCod & "','" & RHConceptoDescuentos.gRHConcDctoPrestAdmin & "'," & pnMontoSoles & ",'" & pcMovNro & "')"
    Else
        sqlD = " Update RHConcepto" _
             & " Set nRHConceptoValor = nRHConceptoValor + " & pnMontoSoles & "" _
             & " Where cRHPlanillaCod = '" & pcPlaCod & "' And cPersCod = '" & pcPersCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdmin & "'"
    End If
    rs.Close
    
    oCon.Ejecutar sqlD
    
    InsertaCreditosAdm = True
    Exit Function
GetFormulaErr:
    InsertaCreditosAdm = False
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902FC00F1
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Sub BeginTrans()
    oCon.BeginTrans
End Sub

Public Sub CommitTrans()
    oCon.CommitTrans
End Sub

Public Sub RollBackTrans()
    oCon.RollBackTrans
End Sub

Public Function GetRHCreditosAdmInst() As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sql = " Select PE.cPersNombre, RH.cPersCod From RHCreditosAdmInstit RH" _
        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
        & " Where RH.nEstado = 1"
    
    Set GetRHCreditosAdmInst = oCon.CargaRecordSet(sql)
End Function

'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function EliminaCreditosAdmOtros(psPlaCod As String, psPeriodo As String) As Boolean
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    
    sqlD = " Delete RHCreditosAdmOtros " _
         & " Where cPlaCod = '" & psPlaCod & "' And cPeriodo = '" & psPeriodo & "'"
    oCon.Ejecutar sqlD
    
    sqlD = " Delete RHConcepto " _
         & " Where cRHPlanillaCod = '" & psPlaCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdminOtros & "'"
    oCon.Ejecutar sqlD
    
    EliminaCreditosAdmOtros = True
    Exit Function
GetFormulaErr:
    EliminaCreditosAdmOtros = False
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function



'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
Public Function InsertaCreditosAdmOtros(pcPersCod As String, pcPlaCod As String, pcPeriodo As String, pcMoneda As String, pnMonto As String, pnTipCambio As String, pcMovNro As String, pcPersCodInst As String, pnMontoSoles As String) As Boolean
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlD = " Insert RHCreditosAdmOtros (cPersCod,cPlaCod,cPeriodo,nMoneda,nMonto,nTipCambio,cMovNro,cPersCodInst)" _
         & " Values('" & pcPersCod & "','" & pcPlaCod & "','" & pcPeriodo & "'," & pcMoneda & "," & pnMonto & "," & pnTipCambio & ",'" & pcMovNro & "','" & pcPersCodInst & "')"
    oCon.Ejecutar sqlD
    
    sqlD = " Select nRHConceptoValor From RHConcepto" _
         & " Where cRHPlanillaCod = '" & pcPlaCod & "' And cPersCod = '" & pcPersCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdminOtros & "'"
    Set rs = oCon.CargaRecordSet(sqlD)
    
    If rs.BOF And rs.EOF Then
        sqlD = " Insert RHConcepto (cRHPlanillaCod,cPersCod,cRHConceptoCod,nRHConceptoValor,cUltimaActualizacion)" _
             & " Values('" & pcPlaCod & "','" & pcPersCod & "','" & RHConceptoDescuentos.gRHConcDctoPrestAdminOtros & "'," & pnMontoSoles & ",'" & pcMovNro & "')"
    Else
        sqlD = " Update RHConcepto" _
             & " Set nRHConceptoValor = nRHConceptoValor + " & pnMontoSoles & "" _
             & " Where cRHPlanillaCod = '" & pcPlaCod & "' And cPersCod = '" & pcPersCod & "' And cRHConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoPrestAdminOtros & "'"
    End If
    rs.Close
    
    oCon.Ejecutar sqlD
    
    InsertaCreditosAdmOtros = True
    Exit Function
GetFormulaErr:
    InsertaCreditosAdmOtros = False
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

