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
'
'Public Function VerificandoDuplicidadDoc(ByVal popersona As DCOMPersona) As Integer
'Dim oPersonas As DCOMPersonas
'Dim i As Integer
'
'    On Error GoTo ErrorVeirficandoDuplicidadDoc
'    Set oPersonas = New DCOMPersonas
'    VerificandoDuplicidadDoc = -1
'    For i = 0 To popersona.NumeroDocumentos - 1
'           If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
'             If oPersonas.BuscaDocumento(CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))), popersona.ObtenerNumeroDoc(i), popersona.PersCodigo) Then
'                VerificandoDuplicidadDoc = i
'                Set oPersonas = Nothing
'                Exit Function
'             End If
'         End If
'    Next i
'    Set oPersonas = Nothing
'    Exit Function
'
'ErrorVeirficandoDuplicidadDoc:
'    Set oPersonas = Nothing
'    Err.Raise Err.Number, "BuscaDocumento", Err.Description
'End Function
'
'Public Function VerificandoTamanioDoc(ByVal popersona As DCOMPersona) As Integer
'Dim i As Integer
'
'    On Error GoTo ErrorVerificandoTamanioDoc
'    VerificandoTamanioDoc = -1
'    For i = 0 To popersona.NumeroDocumentos - 1
'         If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
'            If (Len(popersona.ObtenerNumeroDoc(i)) <> 8 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdDNI) Or (Len(popersona.ObtenerNumeroDoc(i)) <> 11 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdRUC) Then
'               VerificandoTamanioDoc = i
'               Exit Function
'            End If
'        End If
'    Next i
'    Exit Function
'
'ErrorVerificandoTamanioDoc:
'    Err.Raise Err.Number, "BuscaDocumento", Err.Description
'End Function
'
'Public Function VerificandoHomonimia(ByVal popersona As DCOMPersona) As ADODB.Recordset
'Dim oPersonas As DCOMPersonas
'Dim R As ADODB.Recordset
'Dim sNomTmp As String
'
'    On Error GoTo ErrorVerificandoTamanioDoc
'    If popersona.Personeria = 1 Then
'        sNomTmp = popersona.ApellidoPaterno & "/" & popersona.ApellidoMaterno
'    Else
'        sNomTmp = popersona.NombreCompleto
'    End If
'    Set oPersonas = New DCOMPersonas
'    Set R = oPersonas.BuscaCliente(sNomTmp, BusquedaNombre)
'    Set VerificandoHomonimia = R
'    Set oPersonas = Nothing
'    Exit Function
'
'ErrorVerificandoTamanioDoc:
'    Err.Raise Err.Number, "BuscaDocumento", Err.Description
'End Function
'
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
'
'
