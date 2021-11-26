VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCOMPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'variables locales para almacenar los valores de las propiedades
Public sPersNombre As String
Public sPersCod As String
Public dPersNacCreac As Date
Public sPersDireccDomicilio As String
Public sPersTelefono As String
Public sPersPersoneria As String
Public sPersPersoneriaDesc As String
Public sPersIdnroDNI As String
Public sPersIdnroRUC As String
Public sPersIdnroOtro As String
Public sPersNatSexo As String
Public sPerEstado As String
Public sPersOcupa As Integer
Public sPersTipoDoc As String
Public sPersEdad As Integer 'APRI20170622
Public DocsPers As ADODB.Recordset


Public Sub CargaDatos(ByVal psPersCod As String, psPersNombre As String, _
    ByVal pdPersNacCreac As Date, ByVal psPersDireccDomicilio As String, _
    ByVal psPersTelefono As String, ByVal psPersPersoneria As String, _
    ByVal psPersIdnroDNI As String, ByVal psPersIdnroRUC As String, _
    ByVal psPesrIDNroOtro As String, ByVal psPersNatSexo As String, Optional psPersOcupa As String, Optional psPersTipoDoc As String = "1", Optional ByVal psPersEdad As Integer) 'APRI20170622 ADD psPersEdad
Dim Conn As COMConecta.DCOMConecta
Dim ocad As New COMFunciones.FCOMCadenas
Dim R1 As ADODB.Recordset
Dim sSql As String
Dim sPersoneriaDesc As String
Dim sPerNroDocOtro As String 'WIOR 20121112
   Set Conn = New COMConecta.DCOMConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Sub
   End If
   
   sSql = "Select * from Constante where nConsCod = " & Trim(Str(gPersPersoneria)) & " And nConsValor = " & Trim(psPersPersoneria)
   Set R1 = Conn.CargaRecordSet(sSql)
   If Not R1.BOF And Not R1.EOF Then
        sPersoneriaDesc = Trim(R1!cConsDescripcion)
   Else
        sPersoneriaDesc = ""
   End If
   R1.Close
   Set R1 = Nothing
   
   sSql = "Select C.cConsDescripcion  as Tipo, PID.cPersIDNro,PID.cPersIDTpo from PersID PID left join Constante C on convert(int,PID.cPersIDTpo) = C.nConsValor where PID.cPersCod = '" & Trim(psPersCod) & "' And C.nConsCod = " & Trim(Str(gPersIdTipo)) & "  and PID.cPersIDTpo not in (1,2)"  'WIOR 20121112 agrego PID.cPersIDTpo not in (1,2)
   
   Set R1 = Conn.CargaRecordSet(sSql)
   Conn.CierraConexion
   Set Conn = Nothing
   'WIOR 20121112
   If R1.RecordCount > 0 Then
        If Not (R1.EOF And R1.BOF) Then
            sPerNroDocOtro = Trim(R1!cPersIdNro)
        Else
            sPerNroDocOtro = ""
        End If
   Else
    sPerNroDocOtro = ""
   End If
   'WIOR FIN
   
    sPersNombre = ocad.PstaNombre(psPersNombre)
    sPersCod = psPersCod
    dPersNacCreac = pdPersNacCreac
    sPersDireccDomicilio = psPersDireccDomicilio
    sPersTelefono = psPersTelefono
    sPersPersoneria = psPersPersoneria
    sPersPersoneriaDesc = sPersoneriaDesc
    sPersIdnroDNI = psPersIdnroDNI
    sPersIdnroRUC = psPersIdnroRUC
    'madm 20100708
    'sPersIdnroOtro = IIf(psPersIdnroDNI = "", IIf(psPersPersoneria = 1, R1!cPersIdNro, ""), sPersIdnroOtro)'WIOR 20121112 COMENTADO
    sPersIdnroOtro = IIf(psPersIdnroDNI = "", IIf(psPersPersoneria = 1, sPerNroDocOtro, ""), sPersIdnroOtro) 'WIOR 20121112
    sPersTipoDoc = IIf(psPersIdnroDNI = "", IIf(psPersPersoneria = 1, psPersTipoDoc, "2"), "1")
    sPersNatSexo = psPersNatSexo
    sPersOcupa = IIf(psPersOcupa = "", -1, psPersOcupa)
    sPersEdad = psPersEdad 'APRI20170622
    Set DocsPers = R1
End Sub

Public Sub ObtieneClientexCodigo(ByVal psPersCod As String)
Dim oPersona As DCOMPersonas
Dim R As ADODB.Recordset
    
    On Error GoTo ErrorObtieneClientexCodigo
    Set oPersona = New DCOMPersonas
    Set R = oPersona.BuscaCliente(psPersCod, BusquedaCodigo)
    
    If IsNull(R!cPersIDNroDNI) And IsNull(R!cPersIDNroRuc) Then
        MsgBox "El cliente no tiene documento de identidad ni RUC.", vbInformation, "Aviso"
    End If
    
    Call CargaDatos(R!cPersCod, R!cPersNombre, R!dPersNacCreac, R!cPersDireccDomicilio, IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, Trim(IIf(IsNull(R!cPersIDNroDNI), "", R!cPersIDNroDNI)), IIf(IsNull(R!cPersIDNroRuc), "", R!cPersIDNroRuc), IIf(IsNull(R!cPersIdNro), "", R!cPersIdNro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
    R.Close
    Set R = Nothing
    Exit Sub

ErrorObtieneClientexCodigo:
    Err.Raise Err.Number, "RecuperaPersona", Err.Description
End Sub

Public Sub ObtieneConyugeDePersona(ByVal psPersCod As String)
Dim oPersona As DCOMPersonas
Dim R As ADODB.Recordset
Dim sPersCodCony As String


    Set oPersona = New DCOMPersonas
    sPersCodCony = oPersona.BuscaPersCodConyuge(psPersCod)
    sPersCodCony = Trim(sPersCodCony)
    If sPersCodCony <> "" Then
        Set R = oPersona.BuscaCliente(sPersCodCony, BusquedaCodigo)
        Call CargaDatos(R!cPersCod, R!cPersNombre, R!dPersNacCreac, R!cPersDireccDomicilio, IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, Trim(R!cPersIDNroDNI), IIf(IsNull(R!cPersIDNroRuc), "", R!cPersIDNroRuc), IIf(IsNull(R!cPersIdNro), "", R!cPersIdNro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
        R.Close
        Set R = Nothing
    Else
        sPersCod = ""
    End If
    Exit Sub
    
ErrorObtieneClientexCodigo:
    Err.Raise Err.Number, "RecuperaPersona", Err.Description
End Sub

Public Function ObtieneComentario(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
sql = "Select * from PersComentario where cPerscod='" & psPersCod & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set ObtieneComentario = rs
Set Co = Nothing
Set rs = Nothing
End Function

Public Sub ActComentario(ByVal psPersCod As String, ByVal psComentario As String, _
ByVal psFecha As String, ByVal opt As Integer)
Dim sql As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
If opt = 0 Then
    sql = "Update PersComentario " & _
          " Set cComentario='" & psComentario & "' , dFecha='" & psFecha & "'" & _
          " where cPerscod='" & psPersCod & "'"
Else
    sql = " Insert PersComentario  " & _
          " (cPersCod,cComentario,dFecha) values " & _
          " ('" & psPersCod & "','" & psComentario & "','" & psFecha & "')"
End If
Co.AbreConexion
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Function fgVerificaEmpleado(ByVal psPersCod As String) As Boolean
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta

'    lsSQL = " Select Count(cPersCod) as NroEmp " & _
            " From RRHH Where cPerscod ='" & psPersCod & "' And nRHEstado = 201 "
    'PEAC 20080515
    lsSQL = "exec stp_sel_VerificaEmpleado '" & psPersCod & "'"
    
    Co.AbreConexion
    Set lr = Co.CargaRecordSet(lsSQL)
        If lr!NroEmp > 0 Then
            fgVerificaEmpleado = True
        Else
            fgVerificaEmpleado = False
        End If
    Co.CierraConexion
    Set Co = Nothing
    lr.Close
End Function

Public Function fgVerificaEmpleadoVincualdo(ByVal psPersCod As String) As Boolean
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta

    '''*** PEAC 20090530
    'lsSQL = " Select Count(cPersCod) as NroEmp " & _
            " From PersRelaciones Where cPersRelacPersCod  ='" & psPersCod & "' "
    
    lsSQL = " exec stp_sel_VerificaParentescoEmpleado '" & psPersCod & "' "

    Co.AbreConexion
    Set lr = Co.CargaRecordSet(lsSQL)
        If lr!NroEmp > 0 Then
            fgVerificaEmpleadoVincualdo = True
        Else
            fgVerificaEmpleadoVincualdo = False
        End If
    Co.CierraConexion
    Set Co = Nothing
    lr.Close
End Function

'*** PEAC 20090731
Public Function ValidaEnListaNegativa(ByVal pcNumDNI As String, ByVal pcNumRUC As String, Optional ByRef pnCondicion As Integer = 0) As Boolean
Dim Conn As COMConecta.DCOMConecta
Dim ocad As New COMFunciones.FCOMCadenas
Dim R1 As ADODB.Recordset
Dim sSql As String, lnNumDoc As String
Dim lnTipoDoc As Integer

Dim sPersoneriaDesc As String
   Set Conn = New COMConecta.DCOMConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Function
   End If
   
    ValidaEnListaNegativa = False
   
    If Len(Trim(pcNumDNI)) > 0 Then
        lnNumDoc = pcNumDNI
        lnTipoDoc = 1
    ElseIf Len(Trim(pcNumRUC)) > 0 Then
        lnNumDoc = pcNumRUC
        lnTipoDoc = 2
    Else
        Exit Function
    End If
   
   sSql = " exec stp_sel_VerificaPersListaNegativa " & lnTipoDoc & ",'" & lnNumDoc & "','' " 'JACA 20110601 SE AGRE EL 3ER PARAMETRO=''
   
   Set R1 = Conn.CargaRecordSet(sSql)
   If Not (R1.BOF And R1.EOF) Then
        ValidaEnListaNegativa = True
        pnCondicion = R1!nCondicion
   Else
        ValidaEnListaNegativa = False
   End If
   
   R1.Close
   Set R1 = Nothing
   
   Conn.CierraConexion
   Set Conn = Nothing
   
End Function

Public Function ValidaEnListaNegativaCondicion(ByVal pcNumDNI As String, ByVal pcNumRUC As String, Optional ByRef pnCondicion As Integer = 0, Optional ByVal cNombre As String) As Boolean
Dim Conn As COMConecta.DCOMConecta
Dim ocad As New COMFunciones.FCOMCadenas
Dim R1 As ADODB.Recordset
Dim sSql As String, lnNumDoc As String
Dim lnTipoDoc As Integer
Dim lnNombre As String
Dim sPersoneriaDesc As String
   Set Conn = New COMConecta.DCOMConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Function
   End If
   
    ValidaEnListaNegativaCondicion = False
   
    If Len(Trim(pcNumDNI)) > 0 Then
        lnNumDoc = pcNumDNI
        lnTipoDoc = 1
    ElseIf Len(Trim(pcNumRUC)) > 0 Then
        lnNumDoc = pcNumRUC
        lnTipoDoc = 2
    Else
        Exit Function
    End If
    lnNombre = ExtraerApostrofe(cNombre) 'JGPA20181219 ACTA Nº 197-2018-----
   'lnNombre = cNombre 'Comentado x JGPA20181219
   sSql = " exec stp_sel_VerificaPersListaNegativa_Condicion " & lnTipoDoc & ",'" & lnNumDoc & "','" & lnNombre & "' "
   
   Set R1 = Conn.CargaRecordSet(sSql)
   If Not (R1.BOF And R1.EOF) Then
        If R1!nEstado = 0 Then  '   MADM 20101221
            ValidaEnListaNegativaCondicion = True
            pnCondicion = R1!nCondicion
        End If
   Else
        ValidaEnListaNegativaCondicion = False
   End If
   
   R1.Close
   Set R1 = Nothing
   
   Conn.CierraConexion
   Set Conn = Nothing
   
End Function

'WIOR 20120323-INICIO
Public Function ValidaEnListaNegativaCondicionJustificacion(ByVal pcNumDNI As String, ByVal pcNumRUC As String, Optional ByRef pnCondicion As Integer = 0, Optional ByVal cNombre As String) As Boolean
Dim Conn As COMConecta.DCOMConecta
Dim ocad As New COMFunciones.FCOMCadenas
Dim R1 As ADODB.Recordset
Dim sSql As String, lnNumDoc As String
Dim lnTipoDoc As Integer
Dim lnNombre As String
Dim sPersoneriaDesc As String
   Set Conn = New COMConecta.DCOMConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Function
   End If
   
    ValidaEnListaNegativaCondicionJustificacion = False
   
    If Len(Trim(pcNumDNI)) > 0 Then
        lnNumDoc = pcNumDNI
        lnTipoDoc = 1
    ElseIf Len(Trim(pcNumRUC)) > 0 Then
        lnNumDoc = pcNumRUC
        lnTipoDoc = 2
    Else
        Exit Function
    End If
    
   lnNombre = cNombre
   sSql = " exec stp_sel_VerificaPersListaNegativa_CondicionJustificacion " & lnTipoDoc & ",'" & lnNumDoc & "','" & lnNombre & "' "
   
   Set R1 = Conn.CargaRecordSet(sSql)
   If Not (R1.BOF And R1.EOF) Then
        If R1!Tipo > 0 Then
            ValidaEnListaNegativaCondicionJustificacion = True
        Else
            ValidaEnListaNegativaCondicionJustificacion = False
        End If
   Else
        ValidaEnListaNegativaCondicionJustificacion = False
   End If
   
   R1.Close
   Set R1 = Nothing
   
   Conn.CierraConexion
   Set Conn = Nothing
   
End Function

'WIOR - FIN


'*** PEAC 20100514 - MEMO 502-2010-GM-DI/CMACM (VALIDA QUE HAYA MAS DE UN CREDITO PENDIENTE CON UN MISMO ANALISTA Y CLIENTE)
Public Function fgVerificaCredAnalistaCliente(ByVal psPersCod As String, ByVal psCoduser As String) As Boolean
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
    
    lsSQL = " exec stp_sel_VerificaCredAnalistaCliente '" & psPersCod & "','" & psCoduser & "'"

    Co.AbreConexion
    Set lr = Co.CargaRecordSet(lsSQL)
        If lr!nNumCred > 0 Then
            fgVerificaCredAnalistaCliente = True
        Else
            fgVerificaCredAnalistaCliente = False
        End If
    Co.CierraConexion
    Set Co = Nothing
    lr.Close
End Function

'***Agregado por ELRO el 20121031, según OYP-RFC101-2012
Public Function devolverDatosPersona(ByVal pnPersIDTpo As Integer, ByVal psPersIDNro As String) As ADODB.Recordset
Dim oDCOMConecta As New COMConecta.DCOMConecta
Dim lsSQL As String

lsSQL = "exec stp_sel_RFC1012012_DevolverDatosPersona " & pnPersIDTpo & ",'" & psPersIDNro & "'"

oDCOMConecta.AbreConexion
Set devolverDatosPersona = oDCOMConecta.CargaRecordSet(lsSQL)
oDCOMConecta.CierraConexion
Set oDCOMConecta = Nothing
End Function
'***Fin Agregado por ELRO el 20121031*******************

'RECO20150715 ERS032-2014*******************************
Public Function fgVerificaClienteVincEmpleadoSICMAC(ByVal psPersCod As String) As String
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta

    lsSQL = " stp_sel_ObtieneCodigoEmpleadoParentesco'" & psPersCod & "' "

    Co.AbreConexion
    Set lr = Co.CargaRecordSet(lsSQL)
    If Not (lr.EOF And lr.BOF) Then
        fgVerificaClienteVincEmpleadoSICMAC = lr!cPersCodVinc
    End If
    Co.CierraConexion
    Set Co = Nothing
    lr.Close
End Function
'RECO FIN **********************************************
'**********APRI20170630 TI-ERS025-2017******************
Public Function ObtenerVinculadoRiesgoUnico(ByVal psPersCod As String, ByVal psCtaCod As String, ByVal psTipo As Integer) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim Con As COMConecta.DCOMConecta
    Set Con = New COMConecta.DCOMConecta
    
    sSql = "Exec SP_SEL_VINCULADO_RIESGO_UNICO '" & psPersCod & "','" & psCtaCod & "','" & psTipo & "'"
    
    Con.AbreConexion
    Set ObtenerVinculadoRiesgoUnico = Con.CargaRecordSet(sSql)
    Con.CierraConexion
    Set Con = Nothing
End Function
'*********************END APRI*************************
'JGPA20181219 ACTA Nº 197-2018-----------------------
Public Function ExtraerApostrofe(ByVal psCadena As String) As String
Dim sCaracter As String
sCaracter = "'"
If Len(psCadena) > 0 Then
    If InStr(psCadena, sCaracter) > 0 Then
        ExtraerApostrofe = Replace(psCadena, sCaracter, "")
    Else
        ExtraerApostrofe = psCadena
    End If
End If
End Function
'End JGPA------------------------------------------
