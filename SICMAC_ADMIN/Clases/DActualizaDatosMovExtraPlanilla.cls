VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaMovExtraPlanilla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9AF9C70113"
'Actualiza un movimiento extra planilla
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FA0210
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902FA036F
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902FA0333
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

'Agrega un nuevo movimiento extra planilla
'##ModelId=3A9AF9E4022D
Public Function AgregaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaExtPlanillaErr
    
    
    'your code goes here...

    Exit Function
AgregaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:AgregaExtPlanilla Method")
End Function

'Elimina un movimiento extra planilla
'##ModelId=3A9AF9E40269
Public Function EliminaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRHPersCod As String, psCodCta As String) As Boolean
    On Error GoTo EliminaExtPlanillaErr

    'your code goes here...

    Exit Function
EliminaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:EliminaExtPlanilla Method")
End Function

'Actualiza un movimiento extra planilla
'##ModelId=3A9AF9E402A5
Public Function ModificaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String, psMonto As String, psComentario As String, psTipoOpe As String) As Boolean
    On Error GoTo ModificaExtPlanillaErr
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rsE As ADODB.Recordset
    Dim lsCadena As String
    Dim lsMonto As String
    Dim lsComentario As String
    Dim lsTpo As String
    
    Set rsE = New ADODB.Recordset
    
    lsCadena = ""
    While Not prRS.EOF
        If lsCadena = "" Then
            lsCadena = prRS.Fields(0) & prRS.Fields(4) & Right(prRS.Fields(5), 1) & prRS.Fields(10) & Right(prRS.Fields(7), 1) & prRS.Fields(8)
        Else
            lsCadena = lsCadena & "','" & prRS.Fields(0) & prRS.Fields(4) & Right(prRS.Fields(5), 1) & prRS.Fields(10) & Right(prRS.Fields(7), 1) & prRS.Fields(8)
        End If
        prRS.MoveNext
    Wend
    
    sqlE = " Delete RHExtraPlanilla" _
         & " Where cRRHHPeriodo =  '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHPersCod+cCodCta+Rtrim(Ltrim(STR(nTipoOpe)))+cCtaCodRef+Rtrim(Ltrim(STR(nTipRef)))+cPersCodRef Not In ('" & lsCadena & "') "
    prRS.MoveFirst
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        oCon.Ejecutar sqlE
        
        While Not prRS.EOF
            lsMonto = IIf(psMonto = "", prRS.Fields(2), psMonto)
            lsComentario = IIf(psMonto = "", prRS.Fields(3), psComentario)
            lsTpo = IIf(psTipoOpe = "", Right(prRS.Fields(5), 2), psTipoOpe)
            
            
            sqlE = " Select cRRHHPeriodo From RHExtraPlanilla" _
                 & " Where cRRHHPeriodo =  '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHPersCod = '" & prRS.Fields(0) & "' And cCodCta = '" & prRS.Fields(4) & "' And nTipoOpe = " & lsTpo & " And cCtaCodRef = '" & prRS.Fields(10) & "' And nTipRef = " & Right(prRS.Fields(7), 1) & " And cPersCodRef = '" & prRS.Fields(8) & "'"
            Set rsE = oCon.CargaRecordSet(sqlE)
            
            If rsE.EOF And rsE.BOF Then
                sqlE = " Insert RHExtraPlanilla(cRRHHPeriodo, cPlanillaTpoCod, cRHPersCod, cCodCta, nRHExtraMonto, cRHExtraComentario, cUltimaActualizacion,nTipoOpe, nTipRef, cPersCodRef, cCtaCodRef)" _
                     & " Values('" & psRRHHPeriodo & "','" & psPlanillaTpoCod & "','" & prRS.Fields(0) & "','" & prRS.Fields(4) & "'," & lsMonto & ",'" & lsComentario & "','" & psUltimaActualizacion & "'," & lsTpo & "," & Right(prRS.Fields(7), 1) & ",'" & prRS.Fields(8) & "','" & prRS.Fields(10) & "')"
            Else
                sqlE = " Update RHExtraPlanilla" _
                     & " Set nRHExtraMonto = " & lsMonto & " , cRHExtraComentario = '" & lsComentario & "' , cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                     & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHPersCod = '" & prRS.Fields(0) & "' And cCodCta = '" & prRS.Fields(4) & "' And nTipoOpe = " & lsTpo & "  And cCtaCodRef = '" & prRS.Fields(10) & "' And nTipRef = " & Right(prRS.Fields(7), 1) & " And cPersCodRef = '" & prRS.Fields(8) & "'"
            End If
            
            rsE.Close
            oCon.Ejecutar sqlE
            
            prRS.MoveNext
        Wend
        oCon.CommitTrans
        oCon.CierraConexion
    End If
     
    Exit Function
ModificaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:ModificaExtPlanilla Method")
End Function


Public Function MarcaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psTipoOpe As String, psPersCod As String, psCtaCod As String, psDocumento As String, psDocumentoRef As String, psCtaReferencia As String) As Boolean
    On Error GoTo ModificaExtPlanillaErr
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Update RHExtraPlanilla" _
         & " Set cRHExtraDoc = '" & psDocumento & "', cRHExtraDocRef = '" & psDocumentoRef & "' " _
         & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHPersCod = '" & psPersCod & "' And cCodCta = '" & psCtaCod & "' And nTipoOpe = " & psTipoOpe & " And cCtaCodRef = '" & psCtaReferencia & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
     
    Exit Function
ModificaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:ModificaExtPlanilla Method")
End Function


'Devuelve un movimiento extra planilla, de una cuenta
'##ModelId=3A9AFA1A0018
Public Function GetExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, Optional pbConDocumento As Boolean = False, Optional psPersCod As String = "", Optional psCtaCod As String = "") As Recordset
    On Error GoTo GetExtPlanillaErr
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    '
    If psPersCod = "" Then
        If Not pbConDocumento Then
            sql = " Select RHE.cRHPersCod Codigo, PE.cPersNombre Nombre, RHE.nRHExtraMonto Monto, RHE.cRHExtraComentario Comentario, RHE.cCodCta Cta, " _
                & "        Case RHE.nTipoOpe When 0 Then 'CARGOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) Else 'ABONOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) End Tipo, '' A," _
                & "        CO.cConsDescripcion + Space(50) + LTrim(Str(RHE.nTipRef)) Destino, IsNull(RHE.cPersCodRef,'') [Cod Ref], IsNull(PERef.cPersNombre,'') Referencia, IsNull(RHE.cCtaCodRef,'') [Cta Referencia]" _
                & " From RHExtraplanilla RHE" _
                & "      Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & "      Inner Join Constante CO On CO.nConsCod = 6053 And CO.nConsValor = RHE.nTipRef" _
                & "      Left Join Persona PERef On PERef.cPersCod = RHE.cPersCodRef" _
                & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHExtraDoc is Null Order By Nombre"
        Else
            sql = " Select RHE.cRHPersCod Codigo, PE.cPersNombre Nombre, RHE.nRHExtraMonto Monto, RHE.cRHExtraComentario Comentario, RHE.cCodCta Cta, " _
                & "        Case RHE.nTipoOpe When 0 Then 'CARGOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) Else 'ABONOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) End Tipo, '' A," _
                & "        CO.cConsDescripcion + Space(50) + LTrim(Str(RHE.nTipRef)) Destino, IsNull(RHE.cPersCodRef,'') [Cod Ref], IsNull(PERef.cPersNombre,'') Referencia, IsNull(RHE.cCtaCodRef,'') [Cta Referencia], cRHExtraDoc Doc, cRHExtraDocRef DocRef" _
                & " From RHExtraplanilla RHE" _
                & "      Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & "      Inner Join Constante CO On CO.nConsCod = 6053 And CO.nConsValor = RHE.nTipRef" _
                & "      Left Join Persona PERef On PERef.cPersCod = RHE.cPersCodRef" _
                & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' Order By Nombre"
            
            'sql = " Select cRHPersCod Codigo, cPersNombre Nombre, nRHExtraMonto Monto, Case nTipoOpe When 0 Then 'Cargo ' Else 'Abono' End + cRHExtraComentario Comentario, cCodCta Cta , Case nTipoOpe When 0 Then 'CARGOS' + Space(50) + LTrim(Str(nTipoOpe)) Else 'ABONOS' + Space(50) + LTrim(Str(nTipoOpe)) End,'' ,cRHExtraDoc Doc   From RHExtraplanilla RHE " _
                & " Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' Order By cPersNombre"
        End If
    Else
        If Not pbConDocumento Then
            sql = " Select RHE.cRHPersCod Codigo, PE.cPersNombre Nombre, RHE.nRHExtraMonto Monto, RHE.cRHExtraComentario Comentario, RHE.cCodCta Cta, " _
                & "        Case RHE.nTipoOpe When 0 Then 'CARGOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) Else 'ABONOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) End Tipo, '' A," _
                & "        CO.cConsDescripcion + Space(50) + LTrim(Str(RHE.nTipRef)) Destino, IsNull(RHE.cPersCodRef,'') [Cod Ref], IsNull(PERef.cPersNombre,'') Referencia, IsNull(RHE.cCtaCodRef,'') [Cta Referencia]" _
                & " From RHExtraplanilla RHE" _
                & "      Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & "      Inner Join Constante CO On CO.nConsCod = 6053 And CO.nConsValor = RHE.nTipRef" _
                & "      Left Join Persona PERef On PERef.cPersCod = RHE.cPersCodRef" _
                & " where RHE.cRRHHPeriodo = '" & psRRHHPeriodo & "' And RHE.cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And RHE.cRHExtraDoc Is Null" _
                & " And PE.cPersCod = '" & psPersCod & "' And RHE.cCodCta = '" & psCtaCod & "' Order By Nombre"
            
            'sql = " Select cCodCta,nTipoOpe,cRHExtraDoc,nRHExtraMonto , cRHExtraComentario from rhextraplanilla RHE " _
                & " Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cRHExtraDoc is Null And cPersCod = '" & psPersCod & "' And cCodCta = '" & psCtaCod & "' Order By cPersNombre"
        Else
            sql = " Select RHE.cRHPersCod Codigo, PE.cPersNombre Nombre, RHE.nRHExtraMonto Monto, RHE.cRHExtraComentario Comentario, RHE.cCodCta Cta, " _
                & "        Case RHE.nTipoOpe When 0 Then 'CARGOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) Else 'ABONOS' + Space(50) + LTrim(Str(RHE.nTipoOpe)) End Tipo, '' A," _
                & "        CO.cConsDescripcion + Space(50) + LTrim(Str(RHE.nTipRef)) Destino, IsNull(RHE.cPersCodRef,'') [Cod Ref], IsNull(PERef.cPersNombre,'') Referencia, IsNull(RHE.cCtaCodRef,'') [Cta Referencia], cRHExtraDoc Doc, cRHExtraDocRef DocRef" _
                & " From RHExtraplanilla RHE" _
                & "      Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & "      Inner Join Constante CO On CO.nConsCod = 6053 And CO.nConsValor = RHE.nTipRef" _
                & "      Left Join Persona PERef On PERef.cPersCod = RHE.cPersCodRef" _
                & " where RHE.cRRHHPeriodo = '" & psRRHHPeriodo & "' And RHE.cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And RHE.cRHExtraDoc Is Null" _
                & " And PE.cPersCod = '" & psPersCod & "' And RHE.cCodCta = '" & psCtaCod & "' Order By Nombre"
            
'            sql = " Select cRHPersCod Codigo, cPersNombre + '  (' + cCodCta + ')' Nombre, nRHExtraMonto Monto, Case nTipoOpe When 0 Then 'Cargo ' Else 'Abono' End + cRHExtraComentario Comentario, cRHExtraDoc, cCodCta Cta  from rhextraplanilla RHE " _
                & " Inner Join Persona PE On PE.cPersCod = RHE.cRHPersCod" _
                & " where cRRHHPeriodo = '" & psRRHHPeriodo & "' And cPlanillaTpoCod = '" & psPlanillaTpoCod & "' And cPersCod = '" & psPersCod & "' Order By cPersNombre"
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetExtPlanilla = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:GetExtPlanilla Method")
End Function

'Devuelve un movimiento extra planilla, de una cuenta
'##ModelId=3A9AFA1A0018
Public Function GetListaRRHH(psCodPlanilla As String) As Recordset
    On Error GoTo GetExtPlanillaErr
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    
    sql = " Select RH.cPersCod Codigo, cPersNombre + '  (' + cCtaCod + ')' Nombre from RRHH RH" _
        & " Inner Join RHCuentas RC On RH.cPersCod = RC.cPersCod" _
        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
        & " Inner Join RHPlanillaAplicacion PA On RH.nRHEstado = PA.nRHEstado And PA.cRHPlanillaCod = '" & psCodPlanilla & "'" _
        & " Where substring(cCtaCod,6,3)= '" & IIf(psCodPlanilla = gsRHPlanillaCTS, RHEmpleadoCuentasTpo.RHEmpleadoCuentasTpoCTS, RHEmpleadoCuentasTpo.RHEmpleadoCuentasTpoAhorro) & "'" _
        & " Order by cPersNombre"
    
    If oCon.AbreConexion Then
        Set GetListaRRHH = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:GetExtPlanilla Method")
End Function

Public Function GetPersonasExtraPlanillaAdd() As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Select RHE.cPersCod, RHE.cDescripcion + PE.cPersNombre Decrip, Len(RHE.cPersCod) Nivel FROM RHExtraPlanillaDefCtas RHE" _
        & " Inner Join Persona PE On RHE.cPersCod =  PE.cPersCod"
    oCon.AbreConexion
    
    Set GetPersonasExtraPlanillaAdd = oCon.CargaRecordSet(sql)

End Function

Public Function GetCtaPersonasExtraPlanillaAdd(psPersCod As String) As String
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select RHE.cCtaCod FROM RHExtraPlanillaDefCtas RHE Where RHE.cPersCod =  '" & psPersCod & "'"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetCtaPersonasExtraPlanillaAdd = ""
    Else
        GetCtaPersonasExtraPlanillaAdd = rs.Fields(0)
    End If

    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Devuelve un movimiento extra planilla, de una cuenta
'##ModelId=3A9AFA1A0018
Public Function GetRHExtraPlanillaValida(psPlanillaCod As String, psPeriodo As String, pbPorCuentas As Boolean) As Recordset
    On Error GoTo GetExtPlanillaErr
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    
    If Not pbPorCuentas Then
        sql = "Select sum(nRHExtraMonto) Monto, cPersCodRef, IsNull((Select cDescripcion from RHExtraPlanillaDefCtas where cPersCod = cPersCodRef),'PENDIENTE') Destino from rhextraplanilla where cplanillatpocod = '" & psPlanillaCod & "' and crrhhperiodo = '" & psPeriodo & "' group by cPersCodRef"
    Else
        sql = "Select sum(nRHExtraMonto) Monto, cCtaCodRef, IsNull((Select cDescripcion from RHExtraPlanillaDefCtas where cctacod = cCtaCodRef),'PENDIENTE') Destino from rhextraplanilla where cplanillatpocod = '" & psPlanillaCod & "' and crrhhperiodo = '" & psPeriodo & "' group by cCtaCodRef"
    End If
    
    If oCon.AbreConexion Then
        Set GetRHExtraPlanillaValida = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMovExtraPlanilla:GetExtPlanilla Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902FA026A
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
