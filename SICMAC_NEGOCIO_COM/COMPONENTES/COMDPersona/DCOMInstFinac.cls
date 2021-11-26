VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMInstFinac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum TInstFinExistenciaRegistro
    gInstFinExistenciaRegistro = 1
    gInstFinExistenciaReferencia = 2
End Enum
Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String

Dim oImpresora As COMFunciones.FCOMVarImpresion

Public Function CargaCmacs() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCargaInstituciones
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select I.cPersCod, P.cPersNombre, Nivel = 1, C.cConsDescripcion + space(50) + I.cIFTpo as sTipo, C.cConsDescripcion, I.cSubCtaContCod " _
        & "From InstitucionFinanc I JOIN Persona P on I.cPersCod = P.cPersCod " _
        & "Inner JOIN Constante C ON convert(Int,I.cIFTpo) = C.nConsValor " _
        & "WHERE C.nConsCod = '" & COMDConstantes.gCGTipoIF & "' And nConsValor <> '" & COMDConstantes.gCGTipoIF & "' " _
        & "AND Convert(smallint,I.cIFTpo) = 3 Order By P.cPersNombre "
    
    Set CargaCmacs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargaInstituciones:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Instituciones Financieras", Err.Description
End Function

Public Function CargaInstituciones() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCargaInstituciones
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select I.cPersCod,P.cPersNombre, C.cConsDescripcion + space(50) + I.cIFTpo as sTipo, C.cConsDescripcion, I.cSubCtaContCod "
    sSql = sSql & " From InstitucionFinanc I inner join Persona P on I.cPersCod = P.cPersCod "
    sSql = sSql & " Inner join Constante C ON convert(int,I.cIFTpo) = C.nConsValor"
    sSql = sSql & " WHERE C.nConsCod = '" & gCGTipoIF & "' And nConsValor <> '" & gCGTipoIF & "'"
    sSql = sSql & " Order By P.cPersNombre "
    
    Set CargaInstituciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargaInstituciones:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Instituciones Financieras", Err.Description
End Function

Public Sub InsertaNuevaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorInsertaNuevaInstitucion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO InstitucionFinanc(cPersCod,cIFTpo,cSubCtaContCod)"
    sSql = sSql & "VALUES('" & psPersCod & "','" & Right("00" & psTipoInst, 2) & "','" & psSubCta & "')"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaNuevaInstitucion:
    Err.Raise Err.Number, "Nueva Institucion Financiera", Err.Description
End Sub

Public Sub ActualizaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String, ByVal psTipoInstAnt As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorActualizaInstitucion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "UPDATE InstitucionFinanc SET cIFTpo = '" & Right("00" & Trim(psTipoInst), 2) & "', cSubCtaContCod = '" & psSubCta & "'"
    sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psTipoInstAnt), 2) & "'"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaInstitucion:
    Err.Raise Err.Number, "Actualizar Institucion Financiera", Err.Description
End Sub

Public Sub EliminarInstitucion(ByVal psPersCod As String, ByVal psTpoInst As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorActualizaInstitucion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "DELETE InstitucionFinanc "
    sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psTpoInst), 2) & "'"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaInstitucion:
    Err.Raise Err.Number, "Actualizar Institucion Financiera", Err.Description
End Sub
'MADM 20110611 - 20110224 - store
Public Function CargaChequesBancoCred(ByVal psPersCod As String, ByVal pChqEstado As ChequeEstado, Optional ByVal pbIncGirador As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    'By Capi 15042008
    'sSql = "Select * From DocRec Where cPerscod = '" & psPersCod & "' And nEstado in (" & pChqEstado & " ," & gChqEstValorizado & ") and nProducto not in (232,233,234)"
    
''''    If pcPersCodGirador <> "" Then
''''        sSql = " exec stp_sel_CargaChequesBancoCredGirador '" & psPersCod & "','" & pcPersCodGirador & "'"
''''    '*** PEAC 20090323
''''    'sSql = "Select DR.*,nMontoUsadoCh=dbo.fnc_ObtenerMontoUsadoCheque (DR.nTpoDoc,DR.cNroDoc) From DocRec DR Where cPerscod = '" & psPersCod & "' And nEstado in (" & pChqEstado & " ," & gChqEstEnValorizacion & "," & gChqEstValorizado & ") and nProducto not in (232,233,234)"
''''    Else
''''        sSql = " exec stp_sel_CargaChequesBancoCred '" & psPersCod & "'"
''''    End If
''''    '*** FIN PEAC
    
    If pbIncGirador = "" Then
        sSql = " exec stp_sel_CargaChequesBancoCred '" & psPersCod & "'"
    Else
        sSql = " exec stp_sel_CargaChequesBancoCredGirador '" & psPersCod & "','" & pbIncGirador & "'"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaChequesBancoCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'MADM 20110225 - STORE
Public Function RecuperaBancos(Optional ByVal pbTodos As Boolean = False, Optional ByVal pbIncGirador As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'    sSql = "Select I.cPersCod, P.cpersNombre  "
'    sSql = sSql & " From InstitucionFinanc I Inner Join Persona P ON I.cPersCod = P.cPersCod "
'
'    If Not pbTodos Then
'        sSql = sSql & " where I.cIFTpo='01' "
'    End If
'    sSql = sSql & " Order by P.cPersNombre "
    If pbIncGirador = "" Then
        sSql = "exec stp_sel_RecuperaBancos " & IIf(pbTodos = False, 0, 1) & " "
    Else
        sSql = "exec stp_sel_RecuperaBancosGirador  '" & pbIncGirador & "' "
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaBancos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'END MADM
Public Function RecuperaIFinancieraPersCod(Optional ByVal pCombo As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaIFinancieraPersCod
    Set oConecta = New COMConecta.DCOMConecta
    If pCombo Then
        sSql = "Select P.cPersNombre + space(50) +  P.cpersCod  from InstitucionFinanc IFin inner join " & gConsPersona & "Persona P on IFin.cPersCod = P.cpersCod where cIFTpo='05' order by P.cPersNombre"
    Else
        sSql = "select P.cpersCod,P.cPersNombre from InstitucionFinanc IFin inner join " & gConsPersona & "Persona P on IFin.cPersCod = P.cpersCod where cIFTpo='05' order by P.cPersNombre"
    End If
    oConecta.AbreConexion
    Set RecuperaIFinancieraPersCod = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaIFinancieraPersCod:
    Err.Raise Err.Number, "Recupera Institucion Financiera y Codigo de Persona", Err.Description
    
End Function

Public Function ExisteRegistro(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal gOpcExisReg As TInstFinExistenciaRegistro) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    On Error GoTo ErrorExisteInstitucion
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    If gOpcExisReg = gInstFinExistenciaRegistro Then
        sSql = "Select cPersCod,cIFTpo from " & gConsPersona & "InstitucionFinanc "
        sSql = sSql & " Where cPerscod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psIFTpo), 2) & "'"
    End If
    If gOpcExisReg = gInstFinExistenciaReferencia Then
        sSql = "Select cPersCod,cIFTpo from CtaIF "
        sSql = sSql & " Where cPerscod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psIFTpo), 2) & "'"
    End If
    
    Set R = oConec.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteRegistro = True
    Else
        ExisteRegistro = False
    End If
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    Exit Function

ErrorExisteInstitucion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function GetIFSubCuenta(ByVal psPersCod As String, ByVal psIFTpo As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    On Error GoTo ErrorGetIFSubCuenta
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "SELECT cPersCod, cIFTpo,cSubCtaContCod " _
         & "FROM InstitucionFinanc " _
         & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' "
    Set R = oConec.CargaRecordSet(sSql)
    If Not R.RecordCount = 0 Then
        GetIFSubCuenta = R!cSubCtaContCod
    End If
    R.Close
    oConec.CierraConexion
    Exit Function

ErrorGetIFSubCuenta:
        Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function VarInstitucionFinanciera(ByVal sCodCta As String, ByVal nTipoIF As CGTipoIF) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpSql = "SELECT I.cSubCtaContCod AS Campo FROM InstitucionFinanc I JOIN ProductoPersona PP " _
        & "JOIN Producto P ON PP.cCtaCod = P.cCtaCod ON I.cPersCod = PP.cPersCod WHERE " _
        & "P.cCtaCod = '" & sCodCta & "' And PP.nPrdPersRelac = " & gCapRelPersTitular & " And " _
        & "I.cIFTpo = '" & Format$(nTipoIF, "00") & "'"
        
    Set RegCta = New ADODB.Recordset
    RegCta.CursorLocation = adUseClient
    Set RegCta = oCon.CargaRecordSet(tmpSql)
    Set RegCta.ActiveConnection = Nothing
    If (RegCta.BOF Or RegCta.EOF) Then
        VarInstitucionFinanciera = ""
    Else
        VarInstitucionFinanciera = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarInstFinanMov(ByVal pnMovNro As Long, ByVal nTipoIF As CGTipoIF) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpSql = " Select   i.cSubCtaContCod  as Campo " _
        & "    from     movcmac mc   " _
        & "             join InstitucionFinanc i on i.cperscod=mc.cPerscod   " _
        & "             join persona p on p.cperscod = i.cperscod " _
        & "     where nmovnro =" & pnMovNro & " and I.cIFTpo = '" & Format$(nTipoIF, "00") & "'"
        
    Set RegCta = New ADODB.Recordset
    RegCta.CursorLocation = adUseClient
    Set RegCta = oCon.CargaRecordSet(tmpSql)
    Set RegCta.ActiveConnection = Nothing
    If (RegCta.BOF Or RegCta.EOF) Then
        VarInstFinanMov = ""
    Else
        VarInstFinanMov = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function


Private Sub Class_Initialize()
   Dim oImp As COMDConstSistema.DCOMImpresoras
   Dim CIni As COMConecta.DCOMClasIni

    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    Set oImpresora = New COMFunciones.FCOMVarImpresion
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
'    Set oImpresora = Nothing
        
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub

Private Sub Class_Terminate()
    Set oImpresora = Nothing
End Sub

'*** PEAC 20090323
Public Function CargaMontoUsadoCheque(ByVal psNumCheque As String, ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    'sSql = "Select DR.*,nMontoUsadoCh=dbo.fnc_ObtenerMontoUsadoCheque (DR.nTpoDoc,DR.cNroDoc) From DocRec DR Where cPerscod = '" & psPersCod & "' And nEstado in (" & pChqEstado & " ," & gChqEstEnValorizacion & "," & gChqEstValorizado & ") and nProducto not in (232,233,234)"
    sSql = " exec stp_sel_ObtieneMontoUsadoCheque '" & psNumCheque & "'," & pnMoneda
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaMontoUsadoCheque = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'ALPA 20090708****************************************************
Public Function CargaChequeMontoUsado(ByVal pnTipoDoc As Integer, ByVal psNroDoc As String) As Currency 'ADODB.Recordset
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_obtenerMontoxCheque " & pnTipoDoc & ",'" & psNroDoc & "'"
  
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF Or Not rs.BOF Then
        CargaChequeMontoUsado = IIf(IsNull(rs!nMontoUsado), 0, rs!nMontoUsado)
    Else
        CargaChequeMontoUsado = 0
    End If
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*****************************************************************
'MADM 20110425 - 20110224
Public Function ExisteRegistroCheque2(ByVal pcTipo As String, ByVal pcNumero As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    On Error GoTo ExisteRegistroCheque2
    ExisteRegistroCheque2 = False
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    sSql = " exec stp_val_DevolverChequeDuplicado '" & pcTipo & "', '" & pcNumero & "'"
        
    Set R = oConec.CargaRecordSet(sSql)
    If R!valor > 1 Then
        ExisteRegistroCheque2 = True
    End If
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    Exit Function

ExisteRegistroCheque2:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CargaChequeMontoUsadoInstitucion(ByVal pnTipoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, Optional ByVal pcPersCodGirador As String = "") As Currency 'ADODB.Recordset
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_DevolverValorChequeDuplicado " & pnTipoDoc & ",'" & psNroDoc & "','" & psPersCod & "', '" & pcPersCodGirador & "'"
  
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF Or Not rs.BOF Then
        CargaChequeMontoUsadoInstitucion = IIf(IsNull(rs!nMontoUsado), 0, rs!nMontoUsado)
    Else
        CargaChequeMontoUsadoInstitucion = 0
    End If
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*****************************************************************
'MADM 20110320 - STORE girador
Public Function CargaChequesBanco(ByVal psPersCod As String, ByVal pChqEstado As ChequeEstado, Optional ByVal pcPersCodGirador As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    If pcPersCodGirador <> "" Then
        sSql = " exec stp_sel_CargaChequesBancoFiltro '" & psPersCod & "' , " & pChqEstado & " , 1 , '" & pcPersCodGirador & "' "
    Else
        sSql = " exec stp_sel_CargaChequesBancoFiltro '" & psPersCod & "' , " & pChqEstado & " , 0 , '" & pcPersCodGirador & "' "
        'sSql = "Select * From DocRec Where cPerscod = '" & psPersCod & "' And nEstado in (" & pChqEstado & " )"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaChequesBanco = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'MIOL SEGUN RFC098, OBSERVACION EIRE *****************************
Public Function PersoneriaSubasta(ByVal pcPersCod As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    On Error GoTo PersoneriaSubasta
    PersoneriaSubasta = False
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    sSql = " exec stp_val_PersoneriaSubasta '" & pcPersCod & "'"
        
    Set R = oConec.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        PersoneriaSubasta = True
    End If
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    Exit Function

PersoneriaSubasta:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MIOL 20130517, SEGUN RQ13251 **************************************
Public Function CargaCodGirador(ByVal psCodInst As String, ByVal psNumCheque As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_CodGiradorxInstCheque '" & psCodInst & "','" & psNumCheque & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaCodGirador = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function CargaNombreGirador(ByVal psCodInst As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_NombreInstxCod '" & psCodInst & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaNombreGirador = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub ActualizaGiradorCheque(ByVal psPersCod As String, ByVal psNroDoc As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorActualizaGiradorCheque
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Update DocRec Set cPersCodGirador='" & psPersCod & "'"
    sSql = sSql & "Where cNroDoc='" & psNroDoc & "'"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaGiradorCheque:
    Err.Raise Err.Number, "Actualiza Girador", Err.Description
End Sub
'END MIOL **********************************************************
'JUEZ 20140411 *********************************************************
Public Function VerificaEsInstFinanc(ByVal psPersCod As String) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_VerificaEsInstFinanc '" & psPersCod & "'"
    
    VerificaEsInstFinanc = False
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF Or Not rs.BOF Then
        VerificaEsInstFinanc = True
    End If
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END JUEZ **************************************************************
'EJVG 20160702 ERS002-2016
Public Function CargaIFIxCompraDeuda() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = " exec stp_sel_ERS0022016_IFIxCompraDeuda"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaIFIxCompraDeuda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN EJVG 20160702

Public Function ObtieneCtaIFIxCompraDeuda(ByVal psCtaCod As String) As ADODB.Recordset 'LUCV20161212-> Según ERS068-2016
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_ERS0682016_ObtieneCtaIFIxCompraDeuda '" & psCtaCod & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCtaIFIxCompraDeuda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

