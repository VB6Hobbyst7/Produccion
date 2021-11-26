VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

Public Function DigitoChequeo(ByVal psCorrelativo As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 7) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
    liFactor = Array(2, 7, 6, 5, 4, 3, 2)
    liCadena(1) = Val(Mid(psCorrelativo, 1, 1))
    liCadena(2) = Val(Mid(psCorrelativo, 2, 1))
    liCadena(3) = Val(Mid(psCorrelativo, 3, 1))
    liCadena(4) = Val(Mid(psCorrelativo, 4, 1))
    liCadena(5) = Val(Mid(psCorrelativo, 5, 1))
    liCadena(6) = Val(Mid(psCorrelativo, 6, 1))
    liCadena(7) = Val(Mid(psCorrelativo, 7, 1))
    liSum = 0
    For i = 1 To 7
        liSum = liSum + liCadena(i) * liFactor(i - 1)
    Next i
    lnDigito = 11 - (liSum Mod 11)
    If lnDigito = 10 Then
        DigitoChequeo = 0
    ElseIf lnDigito = 11 Then
        DigitoChequeo = 1
    Else
        DigitoChequeo = lnDigito
    End If
End Function
Public Function GenerarNuevoCodigoPersona(ByVal poOpersona As DCOMPersona, ByVal psCodAge As String) As String
Dim lsUltCorr As String
Dim lnCorr As Long
Dim lnCodchequeo As Integer
Dim lsCadFinal As String

    lsUltCorr = poOpersona.UltimoCorrelativo(psCodAge)
    lnCorr = Val(lsUltCorr) + 1
    lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
    lnCodchequeo = DigitoChequeo(lsUltCorr)
    lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
    GenerarNuevoCodigoPersona = Trim(psCodAge) & Trim(lsCadFinal)
End Function

Public Function VerificandoHomonimia(ByVal popersona As DCOMPersona) As ADODB.Recordset
Dim oPersonas As DCOMPersonas
Dim R As ADODB.Recordset
Dim sNomTmp As String

    On Error GoTo ErrorVerificandoTamanioDoc
    If popersona.Personeria = 1 Then
        sNomTmp = popersona.ApellidoPaterno & "/" & popersona.ApellidoMaterno
    Else
        sNomTmp = popersona.NombreCompleto
    End If
    Set oPersonas = New DCOMPersonas
    Set R = oPersonas.BuscaCliente(sNomTmp, BusquedaNombre)
    Set VerificandoHomonimia = R
    Set oPersonas = Nothing
    Exit Function

ErrorVerificandoTamanioDoc:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Public Function VerificandoDuplicidadDoc(ByVal popersona As DCOMPersona) As Integer
Dim oPersonas As DCOMPersonas
Dim i As Integer

    On Error GoTo ErrorVeirficandoDuplicidadDoc
    Set oPersonas = New DCOMPersonas
    VerificandoDuplicidadDoc = -1
    For i = 0 To popersona.NumeroDocumentos - 1
           If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
             If oPersonas.BuscaDocumento(CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))), popersona.ObtenerNumeroDoc(i), popersona.PersCodigo) Then
                VerificandoDuplicidadDoc = i
                Set oPersonas = Nothing
                Exit Function
             End If
         End If
    Next i
    Set oPersonas = Nothing
    Exit Function

ErrorVeirficandoDuplicidadDoc:
    Set oPersonas = Nothing
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Public Function VerificandoTamanioDoc(ByVal popersona As DCOMPersona) As Integer
Dim i As Integer

    On Error GoTo ErrorVerificandoTamanioDoc
    VerificandoTamanioDoc = -1
    For i = 0 To popersona.NumeroDocumentos - 1
         If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
            If (Len(popersona.ObtenerNumeroDoc(i)) <> 8 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdDNI) Or (Len(popersona.ObtenerNumeroDoc(i)) <> 11 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdRUC) Then
               VerificandoTamanioDoc = i
               Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorVerificandoTamanioDoc:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

'Public Sub ChequeoFuenteIngreso(ByRef poOpersona As DCOMPersona, ByVal pnIndice As Integer, ByVal pnTipoFte As PersFteIngresoTipo, ByVal pnPuntDet As Integer)
'Dim dFteFecEval As Date
'Dim DiasLimite As Integer
'Dim dFteFecCaduca As Date
'
'    dFteFecEval = poOpersona.ObtenerFteIngFecEval(pnIndice, pnPuntDet, pnTipoFte)
'    DiasLimite = poOpersona.TiempoFteIngreso
'    dFteFecCaduca = CDate(Format(dFteFecEval, "dd/mm/yyyy")) + DiasLimite
'    Call poOpersona.ActualizarFteIngFecCaduc(dFteFecCaduca, pnIndice, pnPuntDet, pnTipoFte)
'End Sub

'Private Sub CargaUbicacionesGeograficas(Nivel1 As Variant, Nivel2 As Variant, Nivel3 As Variant, Nivel4 As Variant, Nivel5 As Variant, _
'ContNiv1 As Long, ContNiv2 As Long, ContNiv3 As Long, ContNiv4 As Long, ContNiv5 As Long)
'Dim Conn As DCOMConecta
'Dim sSQL As String
'Dim R As ADODB.Recordset
'Dim i As Integer
'
'On Error GoTo ErrCargaUbicacionesGeograficas
'    Set Conn = New DCOMConecta
'    'Carga Niveles
'    sSQL = "select *, 1 p from UbicacionGeografica where cUbiGeoCod like '0%' "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
'    ContNiv1 = 0
'    ContNiv2 = 0
'    ContNiv3 = 0
'    ContNiv4 = 0
'    ContNiv5 = 0
'
'    Conn.AbreConexion
'    Set R = Conn.CargaRecordSet(sSQL)
'    Do While Not R.EOF
'        Select Case R!P
'            Case 1 'Pais
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 2 ' Departamento
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 3 'Provincia
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 4 'Distrito
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 5 'Zona
'                ContNiv5 = ContNiv5 + 1
'                ReDim Preserve Nivel5(ContNiv5)
'                Nivel5(ContNiv5 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'        End Select
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Conn.CierraConexion
'    Set Conn = Nothing
'
'    Exit Sub
'
'ErrCargaUbicacionesGeograficas:
'    MsgBox Err.Description, vbInformation, "Aviso"
'
'End Sub
'Public Sub CargaControles(ByVal pdFecSis As Date, pR_ComboRelPers As ADODB.Recordset, pR_FEDocs As ADODB.Recordset, _
'    pR_RefComercial As ADODB.Recordset, pR_PatVehicular As ADODB.Recordset, pR_RefBancaria As ADODB.Recordset, _
'    MCboTipoSangre As Variant, McmbPersDireccCondicion As Variant, McmbPersPersoneria As Variant, _
'    McmbPersJurMagnitud As Variant, McmbPersNatEstCiv As Variant, _
'    MCmbRela As Variant, MCboPersCiiu As Variant, McmbPersJurTpo As Variant, Nivel1 As Variant, Nivel2 As Variant, Nivel3 As Variant, Nivel4 As Variant, Nivel5 As Variant, ContNiv1 As Long, ContNiv2 As Long, ContNiv3 As Long, ContNiv4 As Long, ContNiv5 As Long)
'
'Dim sSQL As String
'Dim Conn As DCOMConecta
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim oConstante As DCOMConstante
'Dim oCtasIF As DCOMPersona
'Dim MComboRelPers() As String
'
'On Error GoTo ERRORCargaControles
'
'    Set oConstante = New DCOMConstante
'    Set Conn = New DCOMConecta
'    Conn.AbreConexion
'
'    'Carga Combo de Flex de Relaciones de Persona
'    'FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
'    Set pR_ComboRelPers = oConstante.RecuperaConstantes(gPersRelacion)
'
'    'Carga Combo de Documentos
'    Set pR_FEDocs = oConstante.RecuperaConstantes(gPersIdTipo)
'
'    'Carga Combo de Tipos de Referencia Comercial
'    Set pR_RefComercial = oConstante.RecuperaConstantes(3028)
'
'    'Carga Combo de Condicion de Patrimonio Vehicular
'    Set pR_PatVehicular = oConstante.RecuperaConstantes(3029)
'
'    Set oConstante = Nothing
'
'    'Carga TextBuscar Instituciones Financieras
'    Set oCtasIF = New DCOMPersona
'    Set pR_RefBancaria = oCtasIF.GetInstFinancieras("0[123]")
'    Set oCtasIF = Nothing
'
'    'Carga Tipos de Sangre
'    Call CargaComboConstante(gPersTpoSangre, MCboTipoSangre)
'
'    'Carga Condiciones de Domicilio
'    Call CargaComboConstante(gPersCondDomic, McmbPersDireccCondicion)
'
'    'Carga Personeria
'    Call CargaComboConstante(gPersPersoneria, McmbPersPersoneria)
'
'    'Carga Magnitud Empresarial
'    Call CargaComboConstante(gPersJurMagnitud, McmbPersJurMagnitud)
'
'    'Carga Condicion de Domicilio
'    Call CargaComboConstante(gPersCondDomic, McmbPersDireccCondicion)
'
'    'Carga Estado Civil
'    Call CargaComboConstante(gPersEstadoCivil, McmbPersNatEstCiv)
'
'    'Carga Combo Relaciones Con La Institucion
'    Call CargaComboConstante(gPersRelacionInst, MCmbRela)
'
'    'Carga Ubicaciones Geograficas
'    Call CargaUbicacionesGeograficas(Nivel1, Nivel2, Nivel3, Nivel4, Nivel5, ContNiv1, ContNiv2, ContNiv3, ContNiv4, ContNiv5)
'
'
'    sSQL = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
'
'    Set R = Conn.CargaRecordSet(sSQL)
'    ReDim MCboPersCiiu(R.RecordCount)
'    Do While Not R.EOF
'        MCboPersCiiu(R.Bookmark) = Trim(R!cCIIUdescripcion) & Space(100) & Trim(R!cCIIUcod)
'        R.MoveNext
'    Loop
'    R.Close
'
'    'CARGA TIPOS DE PERSONA JURIDICA
'    sSQL = "Select cPersJurTpoCod,cPersJurTpoDesc  from persjurtpo Order by cPersJurTpoDesc"
'    Set R = Conn.CargaRecordSet(sSQL)
'    ReDim McmbPersJurTpo(R.RecordCount)
'    Do While Not R.EOF
'        McmbPersJurTpo(R.Bookmark) = Trim(R!cPersJurTpoDesc) & Space(100) & Trim(R!cPersJurTpoCod)
'        R.MoveNext
'    Loop
'
'    R.Close
'
'    Set R = Nothing
'    Conn.CierraConexion
'    Set Conn = Nothing
'    Exit Sub
'
'
'ERRORCargaControles:
'    MsgBox Err.Description, vbExclamation, "Aviso"
'
'End Sub
'Public Sub CargaControlEstadoPersona(ByVal pnTipoPers As Integer, ByRef pDatos As Variant)
'Dim Conn As DCOMConecta
'Dim sSQL As String
'Dim R As ADODB.Recordset
'Dim Datos() As String
'Dim C As Integer
'    Set Conn = New DCOMConecta
'    Conn.AbreConexion
'
'    'Carga Estados de la Persona
'    sSQL = "Select nConsValor,cConsDescripcion From Constante Where nConsCod = " & Trim(Str(gPersEstado)) & " and nConsValor <> " & Trim(Str(gPersEstado))
'    Set R = Conn.CargaRecordSet(sSQL)
'    ReDim Preserve Datos(0)
'    C = 0
'    Do While Not R.EOF
'        If pnTipoPers = gPersonaNat Then
'            If Len(Trim(R!nConsValor)) = 1 Then
'                C = C + 1
'                ReDim Preserve Datos(C)
'                Datos(C - 1) = Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(Str(R!nConsValor)), 2)
'            End If
'        Else
'            If Len(Trim(R!nConsValor)) > 1 Then
'                C = C + 1
'                ReDim Preserve Datos(C)
'                Datos(C - 1) = Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(R!nConsValor), 2)
'            End If
'        End If
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    Conn.CierraConexion
'    Set Conn = Nothing
'    pDatos = Datos
'    Exit Sub
'
'
'End Sub
'

'*** PEAC 20090718
Public Function ImprimePersListaNegativa() As String

Dim vcondC As String
Dim oImpre As COMFunciones.FCOMImpresion
Set oImpre = New COMFunciones.FCOMImpresion

 '   Dim rsCodHojEval As ADODB.Recordset
    Dim rsListaNeg As ADODB.Recordset
    
'    Dim lnTotPatri As Double
    
'    Dim oDAgencia As COMDConstantes.DCOMAgencias
    
''    Dim oNCredito As COMNCredito.NCOMCredDoc
'    Dim oDCredito As COMDCredito.DCOMCredDoc

    Dim oDPers As New COMDPersona.DCOMPersonas

    Dim sCadImp As String, lcTitulo As String
    
    Dim i As Integer, J As Integer, pos As Integer 'madm 20091112

    Screen.MousePointer = 11
       
 '   Set oDCredito = New COMDCredito.DCOMCredDoc
    Set oDPers = New COMDPersona.DCOMPersonas
 '       Set rsCodHojEval = oDCredito.ObtenerCodHojEval()
        Set rsListaNeg = oDPers.ObtenerPersListaNegativa()
'    Set oDCredito = Nothing
    Set oDPers = Nothing


    'If rsCodHojEval.RecordCount = 0 Then
    If (rsListaNeg.EOF And rsListaNeg.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
'    lnTotPatri = 0
    
    pos = 0 'madm 20091112
    
    sCadImp = ""
        sCadImp = sCadImp & "CAJA MAYNAS" & Chr(10)  ''lnSaltoLinDoc
        sCadImp = sCadImp & "RUC : 20103845328" & Chr(10) & Chr(10) ''lnSaltoLinDoc & lnSaltoLinDoc
        
        lcTitulo = "REGISTRO DE CLIENTES - LISTA NEGATIVA"
        'sCadImp = sCadImp & Space((EspHorxPag - Len(lcTitulo)) / 2 - 18) & lcTitulo & lnSaltoLinDoc
        'sCadImp = sCadImp & lcTitulo & Chr(10) & Chr(10)  'lcTitulo & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.Centra(lcTitulo, 130) & Chr(10)
        sCadImp = sCadImp & String(130, "-") & Chr(10) & Chr(10)   'lnSaltoLinDoc & lnSaltoLinDoc
        
    sCadImp = sCadImp & "ITEM  TIPO PERSONA   DOC.    NUMERO DOC.       NOMBRES/RAZON SOCIAL             JUSTIFICACION               FUENTE" & Chr(10)  'lnSaltoLinDoc
    sCadImp = sCadImp & String(130, "-") & Chr(10) ' lnSaltoLinDoc
    
    vcondC = "" 'madm 20100308
    Dim nChangeSituacion As Integer
    nChangeSituacion = 0
    
    Do While Not rsListaNeg.EOF
        
        If rsListaNeg!cCondicion <> vcondC Then 'madm 20100308
            nChangeSituacion = 1
        End If
                            
            If nChangeSituacion = 1 Then 'madm 20100308
                sCadImp = sCadImp & Chr(10) & Chr(10)
                sCadImp = sCadImp & "CONDICION  : " & rsListaNeg!cCondicion & Chr(10)
                sCadImp = sCadImp & String(130, "-") & Chr(10)
                pos = 0
                nChangeSituacion = 0
            End If
            
                pos = pos + 1 'madm 20091112
                sCadImp = sCadImp & oImpre.ImpreFormat(pos, 5, 0) 'madm 20091112
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cTipoPersona, 12)
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cTipoDoc, 5)
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cNumId, 15)
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cNombre, 30)
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cJustificacion, 25)
                sCadImp = sCadImp & oImpre.ImpreFormat(rsListaNeg!cFuente, 28)
                sCadImp = sCadImp & Chr(10)

        vcondC = rsListaNeg!cCondicion 'madm 20100308
        rsListaNeg.MoveNext
        If rsListaNeg.EOF Then
               Exit Do
        End If
    Loop
    rsListaNeg.Close
    Set rsListaNeg = Nothing
    
    ImprimePersListaNegativa = sCadImp
    Screen.MousePointer = 0

End Function
'ALPA 20091007******************************************************
Public Sub ObtenerDatosGrupoEconomicos(ByRef poDatos As ADODB.Recordset, ByRef poGestionVinculados As ADODB.Recordset, ByRef poGestion As ADODB.Recordset, ByVal pnGrupoEco As Integer, _
Optional ByVal pbBuscaTexto As Boolean = False, Optional ByVal psPersEmpresa As String = "")
'Se agrego el parametro pbBuscaTexto con la finalidad de realizar segun un texto JAME20140304
'Se agrego el parametro  psPersEmpresa con la finalidad de realizar segun un texto JAME20140304
Dim oGrupo As COMDPersona.DCOMGrupoE

Dim sNomTmp As String

    
    On Error GoTo ErrorObtenerDatosGrupoEconomicos
    
    Set oGrupo = New COMDPersona.DCOMGrupoE
        'Set poDatos = oGrupo.ObtenerPersonasXGrupoEcon(pnGrupoEco)
        'JAME20140304 ***********
        If pbBuscaTexto = True Then
            Set poDatos = oGrupo.FiltroBuscarEmpresaGrupoEconomico(psPersEmpresa, pnGrupoEco)
        Else
            Set poDatos = oGrupo.ObtenerPersonasXGrupoEcon(pnGrupoEco)
        End If
        'JAME FIN ******************
    Set oGrupo = Nothing
    Set oGrupo = New COMDPersona.DCOMGrupoE
        Set poGestionVinculados = oGrupo.ObtenerRelacionVinculadosxPersona(pnGrupoEco)
    Set oGrupo = Nothing
    Set oGrupo = New COMDPersona.DCOMGrupoE
        Set poGestion = oGrupo.ObtenerRelacionGestion
    Set oGrupo = Nothing
    Exit Sub

ErrorObtenerDatosGrupoEconomicos:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Sub
'PTI1 AGREGADO 20180817 INFORME N° 002-2017 AC TI/CMACM ******************************************************
'Se dejo de usar  ObtenerDatosGrupoEconomicos POR PTI1 20/08/2018
Public Sub ObtenerDatosSoloGrupoEconomicos(ByRef poDatos As ADODB.Recordset, ByRef poGestionVinculados As ADODB.Recordset, ByRef poGestion As ADODB.Recordset, ByVal pnGrupoEco As Integer, _
Optional ByVal pbBuscaTexto As Boolean = False, Optional ByVal psPersEmpresa As String = "")
Dim oGrupo As COMDPersona.DCOMGrupoE

Dim sNomTmp As String
    
    On Error GoTo ErrorObtenerDatosGrupoEconomicos
    
    Set oGrupo = New COMDPersona.DCOMGrupoE
        If pbBuscaTexto = True Then
            Set poDatos = oGrupo.FiltroBuscarSoloEmpresaXGrupoEconomico(psPersEmpresa, pnGrupoEco)
        Else
            Set poDatos = oGrupo.ObtenerSoloPersonasXGrupoEcon(pnGrupoEco)
        End If
        
  
    Exit Sub

ErrorObtenerDatosGrupoEconomicos:
    Err.Raise Err.Number, "OBTENER DATOS", Err.Description
End Sub
'FIN PTI1 2018082017

'ADD PTI1 20180816 INFORME N° 002 2017 AC TI/CMACM******************************************************

Public Sub ObtenerDatosSoloVinculadoxEmpresa(ByRef poDatos As ADODB.Recordset, ByRef poGestionVinculados As ADODB.Recordset, ByRef poGestion As ADODB.Recordset, ByVal psPersCod As String, ByVal pnGrupoEco As Integer, _
Optional ByVal pbBuscaTexto As Boolean = False, Optional ByVal psPersEmpresa As String = "")
Dim oGrupo As COMDPersona.DCOMGrupoE

Dim sNomTmp As String

    
    On Error GoTo ErrorObtenerDatosGrupoEconomicos
    
    Set oGrupo = New COMDPersona.DCOMGrupoE
       
    Set poDatos = oGrupo.ObtenerSoloVinculadoxEmpresa(pnGrupoEco, psPersCod)
    
    
    Set oGrupo = Nothing
    Set oGrupo = New COMDPersona.DCOMGrupoE
        Set poGestionVinculados = oGrupo.ObtenerRelacionVinculadosxPersona(pnGrupoEco)
    Set oGrupo = Nothing
    Set oGrupo = New COMDPersona.DCOMGrupoE
        Set poGestion = oGrupo.ObtenerRelacionGestion
    Set oGrupo = Nothing
    Exit Sub

ErrorObtenerDatosGrupoEconomicos:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Sub
'FIN AGREGADO PTI1

Public Function GrabarPersGrupoEconomico(nGrupoCod As Integer, sPersCod As String, sPersCodOtro As String, nRepresentanteLegal As Integer, nTipoVinculacion As Integer, nPorcenOtro As Currency, nCargo As Integer, nCargoOtro As Integer, sComentarios As String, nEstado As Integer, sPropiedadDirecta As String, sPropiedadIndirecta As String, nForzar As Integer, dFecha As Date) As Integer
Dim oGrupo As COMDPersona.DCOMGrupoE
Dim oDatos As ADODB.Recordset
Dim sNomTmp As String
Dim nRetorno As Integer

    nRetorno = 1
    On Error GoTo ErrorGrabarPersGrupoEconomico
    
    Set oGrupo = New COMDPersona.DCOMGrupoE
        Set oDatos = oGrupo.GrabarPersGrupoEconomico(nGrupoCod, sPersCod, sPersCodOtro, nRepresentanteLegal, nTipoVinculacion, nPorcenOtro, nCargo, nCargoOtro, sComentarios, nEstado, sPropiedadDirecta, sPropiedadIndirecta, nForzar, dFecha)
    Set oGrupo = Nothing
    
    GrabarPersGrupoEconomico = nRetorno
    Exit Function

ErrorGrabarPersGrupoEconomico:
    Err.Raise Err.Number, "Grabar Grupos Economicos", Err.Description
    GrabarPersGrupoEconomico = -2
End Function
'EJVG20120419
Public Function TienePermisoAdicionarPersonaGrupoEco(ByVal pnGrupoEcoCod As Integer, ByVal psGruposUser As String) As Boolean
Dim oGrupo As COMDPersona.DCOMGrupoE
Dim oDatos As ADODB.Recordset

On Error GoTo ErrorTienePermisoAdicionarPersonaGrupoEco
    Set oGrupo = New COMDPersona.DCOMGrupoE
    Set oDatos = oGrupo.TienePermisoAdicionarPersonaGrupoEco(pnGrupoEcoCod, psGruposUser)
    TienePermisoAdicionarPersonaGrupoEco = oDatos!bPermiso
    Set oGrupo = Nothing
    Set oDatos = Nothing
    Exit Function
ErrorTienePermisoAdicionarPersonaGrupoEco:
    Err.Raise Err.Number, "Err: Tiene Permiso Agregar Persona a Grupo Eco", Err.Description
End Function
Public Function ValidaPersona(ByVal psDato As String, ByRef sDescripcion As String, ByRef sPersDireccion As String, ByRef sPersNroDoc As String) As Boolean

Dim DP As COMDPersona.DCOMPersonas
Dim UP As COMDPersona.UCOMPersona
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
ValidaPersona = False
If psDato = "" Then
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
    Exit Function
End If
Set DP = New COMDPersona.DCOMPersonas
Set UP = New COMDPersona.UCOMPersona
Set rs = DP.BuscaCliente(psDato, BusquedaCodigo)

If Not rs.EOF And Not rs.BOF Then
    ValidaPersona = True
  
    sDescripcion = rs!cPersNombre
    sPersDireccion = rs!cPersDireccDomicilio
  
    If rs!nPersPersoneria = gPersonaNat Then
        sPersNroDoc = rs!cPersIDnroDNI & ""
    Else
        sPersNroDoc = rs!cPersIDnroRUC & ""
    End If
Else
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
End If
rs.Close
Set rs = Nothing
Set DP = Nothing
Set UP = Nothing
End Function
'*********************************************************************************

'***Modificado por ELRO, la fecha 20110721, según documento ACTA Nº 168-2011/TI-D

Public Function mostrarDatosPrincipalesClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarDatosPrincipalesClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarDatosPrincipalesClienteProcesoReforzado = oPersona.recuperarDatosPrincipalesClienteProcesoReforzado(pvTitular)
Set oPersona = Nothing
Exit Function

ErrorMostrarDatosPrincipalesClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarDatosPrincipalesClienteProcesoReforzado", Err.Description

End Function

Public Sub ingresarClienteProcesoReforzado(ByVal pvTitular As String, _
                                           ByVal pvCentroLabor As String, _
                                           ByVal pvFechaRegistro As String, _
                                           ByVal pvPersonaAtencion As String, _
                                           ByVal pvPersonaAutoriza As String, _
                                           ByVal pvTipoVinculo As String)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarClienteProcesoReforzado(pvTitular, pvCentroLabor, pvFechaRegistro, pvPersonaAtencion, pvPersonaAutoriza, pvTipoVinculo)
Set oPersona = Nothing
Exit Sub

ErrorIngresarClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarClienteProcesoReforzado", Err.Description

End Sub

Public Function mostarLlaveClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostarLlaveClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostarLlaveClienteProcesoReforzado = oPersona.recuperarLlaveClienteProcesoReforzado(pvTitular)
Set oPersona = Nothing
Exit Function

ErrorMostarLlaveClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostarLlaveClienteProcesoReforzado", Err.Description

End Function

Public Function listarParentescosClienteReforzado(ByVal pcPariente As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorListarParentescosClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set listarParentescosClienteReforzado = oPersona.recuperarParentescosClienteReforzado(pcPariente)
Set oPersona = Nothing
Exit Function

ErrorListarParentescosClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:ListarParentescosClienteReforzado", Err.Description

End Function

Public Function listarTiposDocumentosClienteReforzado() As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorListarTiposDocumentosClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set listarTiposDocumentosClienteReforzado = oPersona.recuperarTiposDocumentosClienteReforzado()
Set oPersona = Nothing
Exit Function

ErrorListarTiposDocumentosClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:ListarTiposDocumentosClienteReforzado", Err.Description

End Function

Public Function listarPaisesClienteReforzado(ByVal pcUbiGeoCod As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorListarPaisesClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set listarPaisesClienteReforzado = oPersona.recuperarPaisesClienteReforzado(pcUbiGeoCod)
Set oPersona = Nothing
Exit Function

ErrorListarPaisesClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:ListarPaisesClienteReforzado", Err.Description

End Function

Public Function listarOcupacionesClienteReforzado() As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorListarOcupacionesClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set listarOcupacionesClienteReforzado = oPersona.recuperarOcupacionesClienteReforzado()
Set oPersona = Nothing
Exit Function

ErrorListarOcupacionesClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:ListarOcupacionesClienteReforzado", Err.Description

End Function

Public Function listarBancosClienteReforzado() As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorListarBancosClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set listarBancosClienteReforzado = oPersona.recuperarBancosClienteReforzado()
Set oPersona = Nothing
Exit Function

ErrorListarBancosClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:ListarBancosClienteReforzado", Err.Description

End Function

Public Sub ingresarParentescoClienteReforzado(ByVal pnClienteProcesoReforzadoId As Integer, _
                                              ByVal pnRelacionId As Integer, _
                                              pvNombreCompleto As String, _
                                              pvDni As String)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarParentescoClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarParentescoClienteReforzado(pnClienteProcesoReforzadoId, pnRelacionId, pvNombreCompleto, pvDni)
Set oPersona = Nothing
Exit Sub

ErrorIngresarParentescoClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarParentescoClienteReforzado", Err.Description

End Sub

Public Sub ingresarConyugeClienteReforzado(ByVal pnClienteProcesoReforzadoId As Integer, _
                                            ByVal pvNombres As String, _
                                            ByVal pvApellidos As String, _
                                            ByVal pnConsCod As Integer, _
                                            ByVal pnConsValor As Integer, _
                                            ByVal pvNumeroDoi As String, _
                                            ByVal pcUbiGeoCod As String, _
                                            ByVal pvFechaNacimiento As String, _
                                            ByVal pvCentroLaboral As String, _
                                            ByVal pnOcupaciones As String, _
                                            ByVal pmIngresoPromedio As Double)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarConyugeClienteReforzado

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarConyugeClienteReforzado(pnClienteProcesoReforzadoId, _
                                               pvNombres, _
                                               pvApellidos, _
                                               pnConsCod, _
                                               pnConsValor, _
                                               pvNumeroDoi, _
                                               pcUbiGeoCod, _
                                               pvFechaNacimiento, _
                                               pvCentroLaboral, _
                                               pnOcupaciones, _
                                               pmIngresoPromedio)
Set oPersona = Nothing
Exit Sub

ErrorIngresarConyugeClienteReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarConyugeClienteReforzado", Err.Description

End Sub

Public Function mostarLlaveConyuge(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostarLlaveConyuge

Set oPersona = New COMDPersona.DCOMPersona
Set mostarLlaveConyuge = oPersona.recuperarLlaveConyuge(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostarLlaveConyuge:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostarLlaveConyuge", Err.Description

End Function

Public Sub ingresarParentescoConyuge(ByVal pnConyugeId As Integer, _
                                     ByVal pnRelacionId As Integer, _
                                     ByVal pvNombreCompleto As String, _
                                     ByVal pvDni As String)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarParentescoConyuge

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarParentescoConyuge(pnConyugeId, pnRelacionId, pvNombreCompleto, pvDni)
Set oPersona = Nothing
Exit Sub

ErrorIngresarParentescoConyuge:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarParentescoConyuge", Err.Description

End Sub

Public Sub ingresarReferenciaEconomica(ByVal pnClienteProcesoReforzadoId As Integer, _
                                       ByVal pvCargoPublico As String, _
                                       ByVal pvFechaInicio As String, _
                                       ByVal pvFechaCese As String, _
                                       ByVal pvEntidadLabora As String, _
                                       ByVal pdIngreso As Double)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarReferenciaEconomica

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarReferenciaEconomica(pnClienteProcesoReforzadoId, pvCargoPublico, pvFechaInicio, pvFechaCese, pvEntidadLabora, pdIngreso)
Set oPersona = Nothing
Exit Sub

ErrorIngresarReferenciaEconomica:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarReferenciaEconomica", Err.Description

End Sub

Public Sub ingresarReferenciaFinanciera(ByVal pnClienteProcesoReforzadoId As Integer, _
                                        ByVal pvBanco As String, _
                                        ByVal pvTipoProducto As String, _
                                        ByVal pvFuncionarioNegocio As String)
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarReferenciaFinanciera

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarReferenciaFinanciera(pnClienteProcesoReforzadoId, pvBanco, pvTipoProducto, pvFuncionarioNegocio)
Set oPersona = Nothing
Exit Sub

ErrorIngresarReferenciaFinanciera:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarReferenciaFinanciera", Err.Description

End Sub

Public Sub ingresarReferenciaPatrimonial(ByVal pnClienteProcesoReforzadoId As Integer, _
                                         ByVal pvBien As String, _
                                         ByVal pdValor As Double)
                        
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarReferenciaPatrimonial

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarReferenciaPatrimonial(pnClienteProcesoReforzadoId, pvBien, pdValor)
Set oPersona = Nothing
Exit Sub

ErrorIngresarReferenciaPatrimonial:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarReferenciaPatrimonial", Err.Description

End Sub

Public Sub ingresarOtroBien(ByVal pnClienteProcesoReforzadoId As Integer, _
                            ByVal pvBien As String)
                        
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarOtroBien

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarOtroBien(pnClienteProcesoReforzadoId, pvBien)
Set oPersona = Nothing
Exit Sub

ErrorIngresarOtroBien:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarOtroBien", Err.Description

End Sub

Public Sub ingresarPersonaJuridica(ByVal pnClienteProcesoReforzadoId As Integer, _
                                   ByVal pvNombreEmpresa As String, _
                                   ByVal pvRuc As String, _
                                   ByVal pvParticipacion As Integer, _
                                   ByVal pdIngreso As Double)
                        
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarPersonaJuridica

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarPersonaJuridica(pnClienteProcesoReforzadoId, pvNombreEmpresa, pvRuc, pvParticipacion, pdIngreso)
Set oPersona = Nothing
Exit Sub

ErrorIngresarPersonaJuridica:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarPersonaJuridica", Err.Description

End Sub

Public Sub ingresarProveedorCliente(ByVal pnClienteProcesoReforzadoId As Integer, _
                                    ByVal pctipo As String, _
                                    ByVal pvNombreEmpresa As String)
                        
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorIngresarProveedorCliente

Set oPersona = New COMDPersona.DCOMPersona
Call oPersona.registrarProveedorCliente(pnClienteProcesoReforzadoId, pctipo, pvNombreEmpresa)
Set oPersona = Nothing
Exit Sub

ErrorIngresarProveedorCliente:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:IngresarProveedorCliente", Err.Description

End Sub

Public Function mostrarDatosClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarDatosClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarDatosClienteProcesoReforzado = oPersona.recuperarDatosClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarDatosClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarDatosClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarDatosConyugeClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarDatosConyugeClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarDatosConyugeClienteProcesoReforzado = oPersona.recuperarDatosConyugeClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarDatosConyugeClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarDatosConyugeClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarReferenciaEconomicaClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarReferenciaEconomicaClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarReferenciaEconomicaClienteProcesoReforzado = oPersona.recuperarReferenciaEconomicaClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarReferenciaEconomicaClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarReferenciaEconomicaClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarReferenciaFinancieraClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarReferenciaFinancieraClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarReferenciaFinancieraClienteProcesoReforzado = oPersona.recuperarReferenciaFinancieraClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarReferenciaFinancieraClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarReferenciaFinancieraClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarReferenciaPatrimonialClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer, ByVal pcTipoPatrimonio As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarReferenciaPatrimonialClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarReferenciaPatrimonialClienteProcesoReforzado = oPersona.recuperarReferenciaPatrimonialClienteProcesoReforzado(pnClienteProcesoReforzadoId, pcTipoPatrimonio)
Set oPersona = Nothing
Exit Function

ErrorMostrarReferenciaPatrimonialClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarReferenciaPatrimonialClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarPersonaJuridicaClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarPersonaJuridicaClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarPersonaJuridicaClienteProcesoReforzado = oPersona.recuperarPersonaJuridicaClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarPersonaJuridicaClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarPersonaJuridicaClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarProveedorClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer, ByVal pctipo As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarProveedorClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarProveedorClienteProcesoReforzado = oPersona.recuperarProveedorClienteProcesoReforzado(pnClienteProcesoReforzadoId, pctipo)
Set oPersona = Nothing
Exit Function

ErrorMostrarProveedorClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarProveedorClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarParienteClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarParienteClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarParienteClienteProcesoReforzado = oPersona.recuperarParienteClienteProcesoReforzado(pnClienteProcesoReforzadoId)
Set oPersona = Nothing
Exit Function

ErrorMostrarParienteClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarParienteClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarParienteConyugeClienteProcesoReforzado(ByVal pnConyugeId As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarParienteConyugeClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarParienteConyugeClienteProcesoReforzado = oPersona.recuperarParienteConyugeClienteProcesoReforzado(pnConyugeId)
Set oPersona = Nothing
Exit Function

ErrorMostrarParienteConyugeClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarParienteConyugeClienteProcesoReforzado", Err.Description

End Function

Public Function mostrarPersNegativaAutorizacion(ByVal pdfecsol As Date, ByVal psNombre As String, ByVal psUser As String, Optional pbolCond As Boolean = False) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarPersNegativaAutorizacion

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarPersNegativaAutorizacion = oPersona.recuperarPersNegativaAutorizacion(pdfecsol, psNombre, psUser)
Set oPersona = Nothing
Exit Function

ErrorMostrarPersNegativaAutorizacion:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarPersNegativaAutorizacion", Err.Description

End Function

Public Function mostrarPersListaNegativa(ByVal pnTipoDoc As Integer, ByVal pcNumDoc As String, ByVal pnCondicion As Integer) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarPersListaNegativa

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarPersListaNegativa = oPersona.recuperarPersListaNegativa(pnTipoDoc, pcNumDoc, pnCondicion)
Set oPersona = Nothing
Exit Function

ErrorMostrarPersListaNegativa:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarPersListaNegativa", Err.Description

End Function

Public Function mostrarRelacionEntidadClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarRelacionEntidadClienteProcesoReforzado

Set oPersona = New COMDPersona.DCOMPersona
Set mostrarRelacionEntidadClienteProcesoReforzado = oPersona.recuperarRelacionEntidadClienteProcesoReforzado(pvTitular)
Set oPersona = Nothing
Exit Function

ErrorMostrarRelacionEntidadClienteProcesoReforzado:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarRelacionEntidadClienteProcesoReforzado", Err.Description

End Function

'***Fin modificado por ELRO******************************************************
'EJVG 20120322
Public Function NecesitaActualizarDatos(ByVal psPersCod As String, ByVal pdFechaSistema As Date) As Boolean
On Error GoTo ErrorObtenerFechaUltimaActualizacion
    Dim oPersona As COMDPersona.DCOMPersona
    Set oPersona = New COMDPersona.DCOMPersona
    Dim ldFecha As Date
    ldFecha = oPersona.obtenerFechaUltimaActualizacion(psPersCod)
    If DateDiff("D", ldFecha, pdFechaSistema) > 365 Then
        NecesitaActualizarDatos = True
    Else
        NecesitaActualizarDatos = False
    End If
    Set oPersona = Nothing
    Exit Function
ErrorObtenerFechaUltimaActualizacion:
    Set oPersona = Nothing
    Err.Raise Err.Number, "NCOMPersona:obtenerFechaUltimaActualizacion", Err.Description
End Function

'MIOL 20120917, SEGUN RQ12272 **************************************************
Public Function MostrarClientexCuenta(ByVal pcCuenta As String) As ADODB.Recordset 'OK
Dim oPersona As COMDPersona.DCOMPersona

On Error GoTo ErrorMostrarClientexCuenta

Set oPersona = New COMDPersona.DCOMPersona
Set MostrarClientexCuenta = oPersona.recuperarClientexCuenta(pcCuenta)
Set oPersona = Nothing
Exit Function
ErrorMostrarClientexCuenta:
Set oPersona = Nothing
Err.Raise Err.Number, "DCOMPersona:MostrarClientexCuenta", Err.Description
End Function

'OK
Public Sub insClienteBloqueadoxCuenta(ByVal pscMovNro As String, ByVal pscPersCod As String, ByVal pscPersNombre As String, ByVal pscCtaCod As String, psnEstado As Integer, ByVal psdFecBloqueo As String)
    Dim oPersona As COMDPersona.DCOMPersona
    Set oPersona = New COMDPersona.DCOMPersona
    
    Call oPersona.insClienteBloqueadoxCuenta(pscMovNro, pscPersCod, pscPersNombre, pscCtaCod, psnEstado, psdFecBloqueo)
    Set oPersona = Nothing

End Sub

'OK
Public Sub updClienteBloqueadoxSobreGiro(ByVal pscPersCod As String, ByVal pscMovDesloqueo As String)
    Dim oPersona As COMDPersona.DCOMPersona
    Set oPersona = New COMDPersona.DCOMPersona
    
    Call oPersona.updClienteBloqueadoxSobreGiro(pscPersCod, pscMovDesloqueo)
    Set oPersona = Nothing

End Sub
'END MIOL **********************************************************************
'WIOR 20121005 **********************************************
Public Sub InsertaPersonaRealizaOperacion(ByVal pnNroMov As Long, ByVal pnCtaCod As String, ByVal pnTipoCliente As Integer, _
                                            ByVal psPersCod As String, ByVal pnPersDOITipo As Integer, ByVal pnPersDOI As String, _
                                            ByVal psPerNombre As String, ByVal pnTipo As Integer, _
                                            Optional ByVal pcOrigen As String = "", Optional ByVal pnCondicion As Integer = 0) 'WIOR 20121114 AGREGO pcOrigen,pnCondicion

On Error GoTo ErrorPersRealiza
Dim oDPersona As COMDPersona.DCOMPersona
Set oDPersona = New COMDPersona.DCOMPersona
    
Call oDPersona.InsertaPersonaRealizaOperacion(pnNroMov, pnCtaCod, pnTipoCliente, psPersCod, pnPersDOITipo, pnPersDOI, psPerNombre, pnTipo, pcOrigen, pnCondicion) 'WIOR 20121114 AGREGO pcOrigen,pnCondicion
Set oDPersona = Nothing

Exit Sub
ErrorPersRealiza:
    Set oDPersona = Nothing
    Err.Raise Err.Number, "NCOMPersona:InsertaPersonaRealizaOperacion", Err.Description
End Sub
'WIOR FIN **************************************************
'WIOR 20130301 **********************************************
Public Sub InsertaPersonasOperacionGen(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnUsuPresente As Integer, _
                                    ByVal pnOpeEfectivo As Integer, ByVal psOrigen As String, ByVal pcOpeCod As String, _
                                    ByVal pnCondicion As Integer, Optional poListaIntervinientes As Variant = Nothing)
                                    

On Error GoTo ErrorPersRealiza
Dim oDPersona As COMDPersona.DCOMPersona
Dim oPersOpe As PersOpeGeneral
Dim i As Integer

Set oDPersona = New COMDPersona.DCOMPersona

If UBound(poListaIntervinientes) > 0 Then
    Call oDPersona.InsertaPersOperacionGen(pnMovNro, psCtaCod, pnUsuPresente, pnOpeEfectivo, psOrigen, pcOpeCod, pnCondicion)
End If

For i = 0 To UBound(poListaIntervinientes)
    oPersOpe = poListaIntervinientes(i)
    Call oDPersona.InsertaPersOperacionGenDet(pnMovNro, oPersOpe.PersCod, oPersOpe.TpoCliente, oPersOpe.TpoDOI, oPersOpe.NroDOI, oPersOpe.Nombre, oPersOpe.TipoRelacion)
Next i

Set oDPersona = Nothing

Exit Sub
ErrorPersRealiza:
    Set oDPersona = Nothing
    Err.Raise Err.Number, "NCOMPersona:InsertaPersonaRealizaOperacion", Err.Description
End Sub
'WIOR FIN **************************************************
'***Agregado por ELRO el 20130701, RFC1306270002****
Public Function verficarRolPersona(ByVal nPersTipo As Integer, ByVal psPersCod As String) As Boolean
Dim oDPersona As COMDPersona.DCOMPersona
Set oDPersona = New COMDPersona.DCOMPersona
    
verficarRolPersona = oDPersona.devolverRolPersona(nPersTipo, psPersCod)
Set oDPersona = Nothing

Exit Function
ErrorVerficarRolPersona:
    Set oDPersona = Nothing
    Err.Raise Err.Number, "DCOMPersona:verficarRolPersona", Err.Description
End Function
'***Fin Agregado por ELRO el 20130701, RFC1306270002
'JUEZ 20131007 *********************************************
Public Function VerificaCargoMantenimientoCampDatos(ByVal psCargoCod As String) As Boolean
    
On Error GoTo ErrorVerificaCargoMantenimientoCampDatos
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sCargos As String
    Dim sCod() As String
    Dim i As Integer
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sCargos = oGen.LeeConstSistema(gConstSistRHCargoCodMantCampDatos)
    
    sCod = Split(sCargos, ",")
    VerificaCargoMantenimientoCampDatos = False
    
    For i = 0 To UBound(sCod)
        If psCargoCod = sCod(i) Then
            VerificaCargoMantenimientoCampDatos = True
            Exit Function
        End If
    Next i
    Exit Function
ErrorVerificaCargoMantenimientoCampDatos:
    MsgBox Err.Description, vbCritical, "Error"
End Function
'END JUEZ **************************************************

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Public Function validarPersonaMayorEdad(ByVal strCodigoPersona As String, ByVal dFechaSistema As Date) As Boolean

Dim bRespuesta As Boolean
Dim rs As Recordset
Dim oPersona As COMDPersona.DCOMPersonas
Dim dFecNac As Date

On Error GoTo ErrorMostrarClientexCuenta
    Set oPersona = New COMDPersona.DCOMPersonas
    Set rs = New Recordset
    Set rs = oPersona.ObtenerDatosPersona(strCodigoPersona) 'RIRO20140407 ERS017
    
    'RIRO20140407 ERS017 Comentado
    'If Not (rs.EOF And rs.BOF) Then
    '   If (Year(dFechaSistema) - Year(rs!dPersNacCreac)) > 18 Then
    '       bRespuesta = True
    '    Else
    '       If (Year(dFechaSistema) - Year(rs!dPersNacCreac)) = 18 Then
    '           If Month(dFechaSistema) = Month(rs!dPersNacCreac) Then
    '               If Day(dFechaSistema) >= Day(rs!dPersNacCreac) Then
    '                   bRespuesta = True
    '               Else
    '                   bRespuesta = False
    '               End If
    '           ElseIf Month(dFechaSistema) > Month(rs!dPersNacCreac) Then
    '               bRespuesta = True
    '           Else
    '               bRespuesta = False
    '           End If
    '       Else
    '               bRespuesta = False
    '       End If
    '    End If
    '
    'End If
    
    'RIRO20140407 ERS017 ***************************************************************
    If Not rs Is Nothing Then
        If Not (rs.EOF And rs.BOF) Then
            dFecNac = rs!dPersNacCreac
            If DateDiff("d", DateAdd("yyyy", 18, dFecNac), dFechaSistema) >= 0 Then
                bRespuesta = True
            Else
                bRespuesta = False
            End If
        End If
    End If
    'END RIRO **************************************************************************
    
    validarPersonaMayorEdad = bRespuesta
    Set oPersona = Nothing
    Exit Function
ErrorMostrarClientexCuenta:
    Err.Raise Err.Number, "DCOMPersona:validarPersonaMayorEdada", Err.Description
End Function
' *** END RIRO
'ORCR20140314**************************
Public Function ObtenerPersonaRHEstado(ByVal psPersCod As String) As Boolean
    Dim oDCOMPersona As New COMDPersona.DCOMPersona
On Error GoTo ErrorObtenerPersonaRHEstado
    Dim lrDatos As ADODB.Recordset
    Set lrDatos = oDCOMPersona.ObtenerPersonaRHEstado(psPersCod)
    
    If lrDatos.RecordCount > 0 Then
        ObtenerPersonaRHEstado = True
    Else
        ObtenerPersonaRHEstado = False
    End If
    
    Exit Function
ErrorObtenerPersonaRHEstado:
    Err.Raise Err.Number, "NCOMPersona:ObtenerPersonaRHEstado", Err.Description
    ObtenerPersonaRHEstado = False
End Function
'END ORCR20140314**************************

' *** AGREGADO POR JAME 20140303 ,SEGUN ERS167-2013"
Public Function validarDevuelveGrupoEconomicoPersona(ByVal psPersCod As String) As Boolean
    Dim oGrupo As COMDPersona.DCOMGrupoE
    Dim rs As Recordset
    Set oGrupo = New COMDPersona.DCOMGrupoE
    Set rs = New Recordset
    Set rs = oGrupo.DevuelveGrupoEconomicoPersona(psPersCod)
        If Not (rs.EOF And rs.BOF) Then
         validarDevuelveGrupoEconomicoPersona = True
        Else
          validarDevuelveGrupoEconomicoPersona = False
        End If
End Function

Public Function FiltroBuscarEmpresaGrupoEconomico(ByVal poDatos As ADODB.Recordset, ByVal psPersNombre As String, ByVal pnGrupoId As Integer) As ADODB.Recordset
    Dim oGrupo As COMDPersona.DCOMGrupoE
    Set oGrupo = New COMDPersona.DCOMGrupoE
    Set poDatos = oGrupo.FiltroBuscarEmpresaGrupoEconomico(psPersNombre, pnGrupoId)
End Function
' *** END JAME
'FRHU ERS077-2015 20151130
Public Function ValidarOperacionActAutoDatos(ByVal psCodOpe As String) As Boolean
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrorValidarOperacionActAutoDatos
    
    Set rs = oDPersona.ValidarOperacionActAutoDatos(psCodOpe)
    If Not (rs.EOF And rs.BOF) Then
        ValidarOperacionActAutoDatos = True
    Else
        ValidarOperacionActAutoDatos = False
    End If
    Exit Function
ErrorValidarOperacionActAutoDatos:
    Err.Raise Err.Number, "NCOMPersona:ErrorValidarOperacionActAutoDatos", Err.Description
    ValidarOperacionActAutoDatos = False
End Function
Public Function ClienteEsPersonaNatural(ByVal psPersCod As String) As Boolean
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrorClienteEsPersonaNatural
    
    Set rs = oDPersona.ClienteEsPersonaNatural(psPersCod)
    If Not (rs.EOF And rs.BOF) Then
        ClienteEsPersonaNatural = True
    Else
        ClienteEsPersonaNatural = False
    End If
    Exit Function
ErrorClienteEsPersonaNatural:
    Err.Raise Err.Number, "NCOMPersona:ErrorClienteEsPersonaNatural", Err.Description
    ClienteEsPersonaNatural = False
End Function
'Public Function ClienteActualizoAutorizoSusDatos(ByVal psPersCod As String) As Boolean ' comentado pot pti1 ers070-2018 06/12/2018
Public Function ClienteActualizoAutorizoSusDatos(ByVal psPersCod As String) As ADODB.Recordset
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrorClienteActualizoAutorizoSusDatos
    
    Set ClienteActualizoAutorizoSusDatos = oDPersona.ClienteActualizoAutorizoSusDatos(psPersCod)
'    Set rs = oDPersona.ClienteActualizoAutorizoSusDatos(psPersCod) 'comentado por pti1 ers070-2018
'
'    If Not (rs.EOF And rs.BOF) Then
'        ClienteActualizoAutorizoSusDatos = True
'    Else
'        ClienteActualizoAutorizoSusDatos = False
'    End If ' fin comnetado pti1

    Exit Function
ErrorClienteActualizoAutorizoSusDatos:
    Err.Raise Err.Number, "NCOMPersona:ClienteActualizoAutorizoSusDatos", Err.Description
   ' ClienteActualizoAutorizoSusDatos = False comentado por pti1 ers070-2018
End Function
'CREADO POR pti1 ers070-2018 06/12/2018*******************
Public Function codOpeBlackList(ByVal psCodOpe As String, ByVal psCodBlo As String) As Boolean
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo codOpeBlackList
     
    Set rs = oDPersona.codOpeBlackList(psCodOpe, psCodBlo)

    If Not (rs.EOF And rs.BOF) Then
        codOpeBlackList = True
    Else
        codOpeBlackList = False
    End If
    
    Exit Function
codOpeBlackList:
    Err.Raise Err.Number, "NCOMPersona:codOpeBlackList", Err.Description
     codOpeBlackList = False
End Function

Public Function DupDocCred(ByVal CodCue As String) As ADODB.Recordset
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo DupDocCred
     
     Set DupDocCred = oDPersona.DupDocCred(CodCue)

    Exit Function
DupDocCred:
    Err.Raise Err.Number, "NCOMPersona:DupDocCred", Err.Description
End Function
Public Sub imprimeCartAut(ByVal CodUsu As String, ByVal snumcuenta As String)
    Dim oDPersona As COMDPersona.DCOMPersona

    On Error GoTo imprimeCartAut
    
     Set oDPersona = New COMDPersona.DCOMPersona
     
     Call oDPersona.InsertarImpresionCartAut(CodUsu, snumcuenta)

  
    Set oDPersona = Nothing
    Exit Sub
imprimeCartAut:
    Err.Raise Err.Number, "NCOMPersona:imprimeCartAut", Err.Description
End Sub
Public Function ObtenerDistritoAgencia(ByVal gsCodAge As String) As String
    Dim oDPersona As New COMDPersona.DCOMPersona
    On Error GoTo ErrorObtenerDistritoAgencia
    
    ObtenerDistritoAgencia = oDPersona.ObtenerDistritoAgencia(gsCodAge)
    Exit Function
ErrorObtenerDistritoAgencia:
    Err.Raise Err.Number, "NCOMPersona:ObtenerDistritoAgencia", Err.Description
End Function
'FIN CREADO PTI1 ers070-2018*******************************************
Public Function ObtenerDatosParaActAutDeCliente(ByVal psPersCod As String) As ADODB.Recordset
    Dim oDPersona As New COMDPersona.DCOMPersona
    On Error GoTo ErrorObtenerDatosParaActAutDeCliente
    
    Set ObtenerDatosParaActAutDeCliente = oDPersona.ObtenerDatosParaActAutDeCliente(psPersCod)
    Exit Function
ErrorObtenerDatosParaActAutDeCliente:
    Err.Raise Err.Number, "NCOMPersona:ObtenerDatosParaActAutDeCliente", Err.Description
End Function
Public Sub ActualizarPersActAutDatos(ByVal psPersCod As String, ByVal psPersNombre As String, _
                                     ByVal psPersNatSexo As String, ByVal psPersNatEstCiv As String, _
                                     ByVal psNacionalidad As String, ByVal psPersIDTpo As String, _
                                     ByVal psPersIDnro As String, ByVal psPersDireccDomicilio As String, _
                                     ByVal psPersRefDomicilio As String, ByVal psPersDireccUbiGeo As String, _
                                     ByVal psPersCelular As String, ByVal psPersTelefono As String, _
                                     ByVal psEmail As String, ByVal psUltimaActualizacion As String)
On Error GoTo ErrorActualizarPersActAutDatos
    Dim oDPersona As COMDPersona.DCOMPersona
    Set oDPersona = New COMDPersona.DCOMPersona
    
    Call oDPersona.ActualizarPersActAutDatos(psPersCod, psPersNombre, psPersNatSexo, psPersNatEstCiv, psNacionalidad, psPersIDTpo, _
                                         psPersIDnro, psPersDireccDomicilio, psPersRefDomicilio, psPersDireccUbiGeo, _
                                         psPersCelular, psPersTelefono, psEmail, psUltimaActualizacion)
    Set oDPersona = Nothing

Exit Sub
ErrorActualizarPersActAutDatos:
    Set oDPersona = Nothing
    Err.Raise Err.Number, "NCOMPersona:ActualizarPersActAutDatos", Err.Description
End Sub
Public Sub InsertarPersActAutDatos(ByVal psMovNro As String, ByVal pnModifDatosSensible As Integer, ByVal psPersCod As String, _
                                   ByRef pMatDatos() As TActAutDatos, ByVal pnAutorizaUsoDatos As Integer, ByVal pnTipoRegistro As Integer, _
                                   ByVal pnTipoCarga As Integer, Optional ByVal pnMovVistoElec As Integer = 0, Optional ByVal nMedio As Integer = 0, Optional ByVal nCodOpe As String = 0) 'modificado por pti1 10/12/2018  se greagó nmedio ncodope
On Error GoTo ErrorActualizarPersActAutDatos
    Dim oDPersona As COMDPersona.DCOMPersona
    Set oDPersona = New COMDPersona.DCOMPersona
    
    Call oDPersona.InsertarPersActAutDatos(psMovNro, pnModifDatosSensible, psPersCod, pMatDatos(), pnAutorizaUsoDatos, pnTipoRegistro, pnTipoCarga, pnMovVistoElec, nMedio, nCodOpe) 'modificado por pti1 10/12/2018  se greagó nmedio, ncodope
    Set oDPersona = Nothing
Exit Sub
ErrorActualizarPersActAutDatos:
    Set oDPersona = Nothing
    Err.Raise Err.Number, "NCOMPersona:ActualizarPersActAutDatos", Err.Description
End Sub
'FRHU 20151210 TIC1512100001
Public Function ValidarCargoParaActAutoDatos(ByVal psCodCargoUser As String) As Boolean
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrorValidarCargoParaActAutoDatos
    
    Set rs = oDPersona.ValidarCargoParaActAutoDatos(psCodCargoUser)
    If Not (rs.EOF And rs.BOF) Then
        ValidarCargoParaActAutoDatos = True
    Else
        ValidarCargoParaActAutoDatos = False
    End If
    Exit Function
ErrorValidarCargoParaActAutoDatos:
    Err.Raise Err.Number, "NCOMPersona:ErrorValidarCargoParaActAutoDatos", Err.Description
    ValidarCargoParaActAutoDatos = False
End Function
'FIN FRHU 20151210
'FIN FRHU 20151130
'FIN FRHU 20151210
'FIN FRHU 20151130
'***JGPA20191210 ACTA N° 106 - 2019
Public Function VerificaPersonaListaNegativa(ByRef psCondicion As String, ByRef psParMensaje As String, ByVal psNombre As String, ByVal psApePat As String, ByVal psDoc As String, ByVal pnTipoBusqueda As Integer, Optional ByVal psApeMat As String = "", Optional ByVal psApeCas As String = "") As Integer
    Dim oDPersonas As New COMDPersona.DCOMPersonas
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrVerificaPersonaListaNegativa
    Set rs = oDPersonas.VerificaPersonaListaNegativa(psNombre, psApePat, psApeMat, psApeCas, psDoc, pnTipoBusqueda) 'JGPA20200117 Added [psDoc,pnTipoBusqueda]
    If Not (rs.BOF And rs.EOF) Then
    
        '20210116LARI AGREGADO POR LARI******************************
        'SI RECUPERA DATOS VERIFICA SI EL NUM DOC DEL DATO RECUPERDO
        'ES IGUAL AL DATO INGRESADO
        'SI NO RECUPERA DATOS PASA CON NORMALIDAD
        Dim NumDocId As String
        NumDocId = Trim(rs!cNumId)
        If ((NumDocId <> Null Or NumDocId <> "") And NumDocId = psDoc) Then
            VerificaPersonaListaNegativa = rs!nCondicion
            psCondicion = Trim(rs!cCondicion)
            psParMensaje = Trim(rs!cParMensaje)
            
        ElseIf ((NumDocId = Null Or NumDocId = "")) Then
            VerificaPersonaListaNegativa = rs!nCondicion
            psCondicion = Trim(rs!cCondicion)
            psParMensaje = Trim(rs!cParMensaje)
            
        Else
            VerificaPersonaListaNegativa = 0
            psParMensaje = ""
        End If
        '************************************************************

        '20210116LARI comendato por LARI
'        VerificaPersonaListaNegativa = rs!nCondicion
'        psCondicion = Trim(rs!cCondicion)
'        psParMensaje = Trim(rs!cParMensaje)
    Else
        VerificaPersonaListaNegativa = 0
        psParMensaje = ""
    End If
    Set oDPersonas = Nothing
    Set rs = Nothing
    Exit Function
ErrVerificaPersonaListaNegativa:
    Err.Raise Err.Number, "Verifica Persona Lista Negativa", Err.Description
End Function
'***End JGPA20191210

