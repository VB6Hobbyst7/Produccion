VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMRCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oCon As COMConecta.DCOMConecta
Dim oImpre As COMFunciones.FCOMImpresion
Dim oCad As COMFunciones.FCOMCadenas
'Public Event ShowProgress()
'Public Event CloseProgress()
'Public Event Progress(pnValor As Long, pnTotal As Long)
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Private Sub Class_Initialize()
Set oCon = New COMConecta.DCOMConecta
Set oImpre = New COMFunciones.FCOMImpresion
Set oCad = New COMFunciones.FCOMCadenas

If oCon.AbreConexion = False Then
    Err.Raise Err.Number, "Error en Conexión", Err.Description
End If
End Sub

Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
Set oImpre = Nothing
Set oCad = Nothing

End Sub

Public Function nCargaParametroRCD(ByVal psFecha As String, ByVal psServConsol As String) As ADODB.Recordset
Dim lsSQL As String
Dim lrs As ADODB.Recordset
lsSQL = "Select * from " & psServConsol & "RCDParametro Where cMes ='" & psFecha & "' "
Set lrs = oCon.CargaRecordSet(lsSQL)

Set nCargaParametroRCD = lrs
End Function

' ** Crea las Tabla de Trabajo del RCD
'Public Sub nCreaTablasRCD(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Boolean = False) JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Public Sub nCreaTablasRCD(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Integer)
'JUEZ 20150310 Se agregó pbRCDTransf
Dim lsSQL As String
Dim lsNomTabla As String 'JUEZ 20150310
'*****************END Comentado by NAGL 20200703******************'
'******Agregado by NAGL 20200703************'
'lsSQL = "EXEC stp_ins_CrearTablasRCD '" & psFecha & "','" & psServConsol & "', " & IIf(pbRCDTransf, 1, 0) & " " JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
lsSQL = "EXEC stp_ins_CrearTablasRCD '" & psFecha & "','" & psServConsol & "', " & pbRCDTransf & " "
oCon.Ejecutar (lsSQL)
'*************END NAGL 20200703*************'
End Sub

'** Verifica si existen las tablas
Public Function nExisteTabla(ByVal lsNombreTabla As String, ByVal psServConsol As String) As Boolean
Dim rs As New ADODB.Recordset
Dim lsSQL As String

nExisteTabla = False
lsSQL = "Select * From " & psServConsol & "SysObjects Where Name = '" & lsNombreTabla & "' and type ='U'"
Set rs = oCon.CargaRecordSet(lsSQL)

If Not (rs.BOF And rs.EOF) Then
    nExisteTabla = True
End If
rs.Close
Set rs = Nothing
End Function

Public Function nObtieneCalificacionPersonaProcesada(ByVal psPersona As String, ByVal psServConsol As String, _
                                                    ByRef psMensaje As String) As String
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim lsCalGen As String
lsCalGen = "0"
lsSQL = "Select cCalGen FROM ColocCalifProv WHERE cPersCod ='" & Trim(psPersona) & "'"
Set rs = oCon.CargaRecordSet(lsSQL)
    
    If rs.BOF And rs.EOF Then
        psMensaje = "No se encontro Calificacion de Persona"
    Else
        If IsNull(rs!cCalGen) Then
            psMensaje = "No se ha asignado calificacion Correcta"
        Else
            lsCalGen = Trim(rs!cCalGen)
        End If
    End If
    rs.Close
    Set rs = Nothing
    
nObtieneCalificacionPersonaProcesada = lsCalGen
End Function

Public Function FormatoCtaContable(ByVal pCuentaCnt As String) As String
  FormatoCtaContable = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function

Public Function EmiteCodigoPersona(lsCodAux As String) As String
Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = "Select * from RCDCodigoAux where cCodAux='" & Trim(lsCodAux) & "'"
Set rs = oCon.CargaRecordSet(Sql)

If Not rs.EOF Then
    EmiteCodigoPersona = Trim(rs!cCodpers)
Else
    EmiteCodigoPersona = ""
End If
rs.Close
Set rs = Nothing
End Function

Private Function EmiteCalificacion(lsCodPers As String) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = "Select * from ColocevalCalif where cperscod='" & lsCodPers & "'"

Set rs = oCon.CargaRecordSet(Sql)

EmiteCalificacion = ""
If Not rs.EOF Then
    EmiteCalificacion = Trim(rs!cEvalCalif)
End If
rs.Close
Set rs = Nothing
End Function

'----------------------------------------
'***********************************************************
'************ Genera los Datos para el RCD *****************
'***********************************************************
Public Sub GeneraDatosRCD(ByVal psFecha As String, ByVal psServConsol As String, _
                        ByVal pdFecDataFM As Date, ByVal pnTipoCambio As Double, _
                        ByVal psCodOfiInf As String, ByVal psUbicaGeoRCD As String, _
                        ByVal pnMontoMinimoRCD As Double, ByVal psCtaNoPref As String, _
                        ByRef psMensaje As Variant, ByRef psMensaje1 As Variant, _
                        Optional ByVal pbRCDTransf As Boolean = False)
                        'JUEZ 20150310 Se agregó pbRCDTransf

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim rsPers As ADODB.Recordset
Dim rsRCD As ADODB.Recordset
Dim nLogDeven As Integer
Dim lbEncuentroEnRCD  As Boolean
Dim loRCDProceso As COMNCredito.NCOMRCD
Dim lsNomTabla As String 'JUEZ 20150310

Dim lsRelacionInst  As String
Dim rsCred As New ADODB.Recordset

Dim i As Long, J As Long
Dim Agencia As String
Dim lscadena As String
Dim lsNombrePersona As String
Dim lsTipoDocRCD As String * 1
Dim lsTipoInfRCD As String * 1
Dim lsCodPers As String
Dim lsIndRCC  As String

Dim lsCredito As String
Dim lnSaldoCap As Currency
Dim lnCapVenc As Currency
Dim lnDiasAtraso As Integer
Dim lbCobJud As Boolean
Dim lbDemanda As Boolean
Dim lbCastigado As Boolean
Dim lbRefinan As Boolean
Dim lnIntDeveng As Currency
Dim lnIntDevengCC As Currency
Dim nSaldoIntSuspensoDeven As Currency
Dim lnProvision As Currency
Dim lnMontoGar1 As Currency  ' Garant Hipotecaria
Dim lnMontoGar2 As Currency  ' Otras Garantias
Dim lnSaldoGasto As Currency 'JUEZ 20121002
Dim lnMontoReprogDesNat As Currency 'EJVG20121212
Dim lnMontoReprogZonaAfectInunda As Currency 'LUCV20170510

Dim lsEstadoCredito As String
Dim lsCuentaCnt As String
    
Dim rsGarant As ADODB.Recordset
    
'Datos usados a Nivel de Formulario
Dim lcCodUnico As String * 15  ' Codigo asignado por la empresa
Dim lcCodSBS As String * 10
Dim lcNomPers As String * 120
Dim lcNomPersComp As String * 120
Dim lcGenero As String * 1
Dim lcEstadoCiv As String * 1
Dim lcActEcon As String * 4
Dim lcCodRegPub As String * 15
Dim lcTiDoTr As String * 1
Dim lcNuDoTr As String * 11
Dim lcTiDoCi As String * 1
Dim lcNuDoCi As String * 12
Dim lcTipPers As String * 1
Dim lcResid As String * 1
Dim lcMagEmp As String * 1
Dim lcAccionista As String * 1
Dim lcRelInst As String * 1
Dim lcPaisNac As String * 4
Dim lcSiglas As String * 20
Dim lcCalificacion As String * 1
Dim lcIndRCambiario As String * 1
Dim lcIndAtraso As String * 1
Dim lcIndClasifInterna As String * 5
Dim lcCondEndeuda As String * 1

Dim lsPat  As String
Dim lsMat  As String
Dim lsCas  As String
Dim lsNom1 As String
Dim lsNom2 As String
'By Capi 0711207
Dim lnMinimo As String
Dim lnAnno As Long
Dim lnMes As Long

lnAnno = Val(Mid(psFecha, 1, 4))
lnMes = Val(Mid(psFecha, 5, 2))

Dim rsSaldosCont As New ADODB.Recordset
Dim lsCondEspecial As String
Dim lsOpeEnGarant As String
Dim lnDiasGarantia As Integer ' Dias desde ult TAsac Garant

Dim gmAgSBS() As String 'EJVG20160709
Dim gmAgUbicacionSBS() As String 'EJVG20160709
    'Obtiene equivalencias de Agencias
    lsSQL = "Select cAgecod, cAgenciaSBS, cUbiGeoSBS, cAgeCodMax = (select max(cAgeCod) from Agencias) from agencias" 'EJVG20160709
    Set rs = oCon.CargaRecordSet(lsSQL)

    Do While Not rs.EOF
        'EJVG20160709 ***
        If rs.Bookmark = 1 Then
            ReDim gmAgSBS(rs!cAgeCodMax)
            ReDim gmAgUbicacionSBS(rs!cAgeCodMax)
        End If
        'END EJVG *******
        
        gmAgSBS(rs!cAgeCod) = rs!cAgenciaSBS
        gmAgUbicacionSBS(rs!cAgeCod) = rs!cUbiGeoSBS
        rs.MoveNext
    Loop
    rs.Close
    
    If pbRCDTransf Then
        lsSQL = "exec stp_sel_RecuperaDatosParaGenerarRCDT " & pnMontoMinimoRCD & ",'" & Format(pdFecDataFM, "yyyy/mm/dd") & "'"
        lsNomTabla = "RCDTvc"
    Else
        lsSQL = "exec stp_sel_RecuperaDatosParaGenerarRCD " & pnMontoMinimoRCD & ",'" & Format(pdFecDataFM, "yyyy/mm/dd") & "'"
        lsNomTabla = "RCDvc"
    End If

    i = 0
    lsTipoDocRCD = "1"
    lsTipoInfRCD = "1"
    lnMinimo = "1"
    Set rs = oCon.CargaRecordSet(lsSQL)
     
    ReDim psMensaje(0)
    ReDim psMensaje1(0)
 
    If Not (rs.BOF And rs.EOF) Then
        '***** LUCV20170411, Modificó y Comentó *****
        lsSQL = "select CG.cCtaCod,rtrim(CG.cTipoGarant) as cTipoGarant,CG.nMoneda, "
        lsSQL = lsSQL & "   sum(round(CG.nMonto,2)) nMonGarantia,"
        lsSQL = lsSQL & " case "
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40201%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40206%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40229%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_4040102%' then left(CG.cCtaContCod,10)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_410' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_41001' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_409%' then left(CG.cCtaContCod,6)"
        lsSQL = lsSQL & "  Else CG.cCtaContCod "
        lsSQL = lsSQL & " end cCtaContCod  into " & psServConsol & "#TmpGarantiasRCD "
        lsSQL = lsSQL & "from " & psServConsol & "CredGarantiasConsol CG  inner join " & psServConsol & "ProductoPersonaConsol PP on CG.cCtaCod = PP.cCtaCod "
        lsSQL = lsSQL & "   inner join " & psServConsol & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20 "
        lsSQL = lsSQL & "where datediff(d,CG.dFecha,'" & Format(pdFecDataFM, "yyyymmdd") & "') = 0 "
        lsSQL = lsSQL & "group bY CG.cCtaCod,CG.cTipoGarant,CG.nMoneda "
        lsSQL = lsSQL & "   ,case "
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40201%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40206%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40229%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_4040102%' then left(CG.cCtaContCod,10)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_410' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_41001' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_409%' then left(CG.cCtaContCod,6)"
        lsSQL = lsSQL & "  Else CG.cCtaContCod "
        lsSQL = lsSQL & " end "
        oCon.Ejecutar (lsSQL)
    
        '************************************************************************************
        lsSQL = " select CG.cCtaCod,rtrim(CG.cTipoGarant) as cTipoGarant,CG.nMoneda,"
        lsSQL = lsSQL & "   sum(round(CG.nMonto,2)) nMonGarantia,"
        lsSQL = lsSQL & " case "
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40201%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40206%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40229%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_4040102%' then left(CG.cCtaContCod,10)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_410' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_41001' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_409%' then left(CG.cCtaContCod,6)"
        lsSQL = lsSQL & "  Else CG.cCtaContCod "
        lsSQL = lsSQL & " end cCtaContCod into " & psServConsol & "#TmpGarantiasRCDCF "
        lsSQL = lsSQL & " from DBConsolidada..CredGarantiasConsol CG  inner join DBConsolidada..ProductoPersonaConsol PP on CG.cCtaCod = PP.cCtaCod"
        lsSQL = lsSQL & " inner join DBConsolidada..CartaFianzaConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20"
        lsSQL = lsSQL & " where datediff(d,CG.dFecha,'" & Format(pdFecDataFM, "yyyymmdd") & "') = 0"
        lsSQL = lsSQL & "     and cFF=''"
        lsSQL = lsSQL & " group bY CG.cCtaCod,CG.cTipoGarant,CG.nMoneda"
        lsSQL = lsSQL & "   ,case "
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40201%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40206%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_40229%' then left(CG.cCtaContCod,8)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_4040102%' then left(CG.cCtaContCod,10)"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_410' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_41001' then left(CG.cCtaContCod,6)+'0000'"
        lsSQL = lsSQL & "    when CG.cCtaContCod  like '84_409%' then left(CG.cCtaContCod,6)"
        lsSQL = lsSQL & "  Else CG.cCtaContCod "
        lsSQL = lsSQL & " end "
        oCon.Ejecutar (lsSQL)
        '***** Fin Comentario LUCV20170411 *****
        '************************************************************************************
        
        Do While Not rs.EOF
            i = i + 1
            '***********************************************
            '**  DATOS DE LA PERSONA  -  RCDvcAAAAMM01   ***
            '***********************************************
            '***** limpio las variables
            lcCodUnico = "":         lcCodSBS = ""
            lcNomPers = "":          lcActEcon = ""
            lcNomPersComp = ""
            lcCodRegPub = "":        lcTiDoTr = ""
            lcNuDoTr = "":           lcTiDoCi = ""
            lcNuDoCi = "":           lcTipPers = ""
            lcResid = "":            lcMagEmp = ""
            lcAccionista = "":       lcRelInst = ""
            lcPaisNac = "":          lcSiglas = ""
            lcCalificacion = "0":    lcCondEndeuda = "" 'JUEZ 20130809 lcCondEndeuda
            lcGenero = ""
            lcEstadoCiv = ""
            lcIndRCambiario = ""
            lcIndAtraso = ""
            lcIndClasifInterna = ""
            
            lsCodPers = Trim(EmiteCodigoAux(Trim(rs!cPersCod), psServConsol))
            If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs!cPersCod)
            End If
            'By Capi 07112007
            lbEncuentroEnRCD = False
            If lsCodPers > lnMinimo Then
                lbEncuentroEnRCD = True
                lnMinimo = lsCodPers
            End If
            'End By
                
            If lbEncuentroEnRCD = True Then 'Lo inserta
                If Not (IsNull(rs!cCodUnico)) Then ' Existe en Maestro Persona
                    'Asigno los valores a la Variables de RCDMaestroPersona
                    lcCodUnico = rs!cCodUnico
                    lcCodSBS = IIf(IsNull(rs!cCodSBS), "", rs!cCodSBS)
                    lcNomPers = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
                    lcNomPers = fReemplazaCaracterEspecial(lcNomPers)
                    lcActEcon = IIf(IsNull(rs!cActEcon), "", rs!cActEcon)
                    lcCodRegPub = IIf(IsNull(rs!cCodRegPub), "", rs!cCodRegPub)
                    lcTiDoTr = IIf(IsNull(rs!cTidotr), "", rs!cTidotr)
                    lcNuDoTr = IIf(IsNull(rs!cNudotr), "", rs!cNudotr)
                    lcTiDoCi = IIf(IsNull(rs!cTidoci), "", rs!cTidoci)
                    lcNuDoCi = IIf(IsNull(rs!cNudoci), "", rs!cNudoci)
                    lcTipPers = rs!cTipPers
                    If Val(lcTipPers) >= "3" Then
                       lcTipPers = "2"
                    End If
                    lcResid = rs!cResid
                    lcMagEmp = IIf(IsNull(rs!cMagEmp), "", rs!cMagEmp)
                    lcAccionista = rs!cAccionista
                    lcRelInst = rs!cRelInst
                    lcPaisNac = rs!cPaisNac
                    lcSiglas = IIf(IsNull(rs!cSiglas), "", rs!cSiglas)
                    lcGenero = IIf(IsNull(rs!cGenero), "", rs!cGenero)
                    lcEstadoCiv = IIf(IsNull(rs!cEstadoCiv), "", rs!cEstadoCiv)
                    
                    lsPat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApePat), "", rs!cApePat)))
                    lsMat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeMat), "", rs!cApeMat)))
                    lsCas = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeCasada), "", rs!cApeCasada)))
                    lsNom1 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cPriNom), "", rs!cPriNom)))
                    lsNom2 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cSegNom), "", rs!cSegNom)))
                Else ' Busco en Personas <>  2 order by cPersIDTpo ASC
                    lsSQL = "exec stp_sel_RecuperaDatosCliente '" & lsCodPers & "'" '***** LUCV20170411, Agregó
                    Set rsPers = oCon.CargaRecordSet(lsSQL)
    
                    If rsPers.BOF And rsPers.EOF Then
                        MsgBox "NO ENCUENTRO LA PERSONA", vbInformation, "Aviso"
                    Else
                        lcCodSBS = rsPers!cCodSBS '"" 'ARCV 16-06-2007
                        'lcNomPersComp = rsPers!cPersNombre
                        lcNomPersComp = fReemplazaCaracterEspecial(rsPers!cPersNombre) 'LUCV20181211, Corrección de apostrofe
                        lcNomPers = fReemplazaCaracterEspecial(rsPers!cPersNombre)
                        CambiaNombreRCD fReemplazaCaracterEspecial(rsPers!cPersNombre), lsPat, lsMat, lsCas, lsNom1, lsNom2 'LUCV20181211: rsPers!cPersNombre
    
                        lcActEcon = Right(rsPers!cActEcon, 4)
                        lcTipPers = rsPers!cTipPers
                        lcMagEmp = rsPers!cMagEmp
                        lcRelInst = rsPers!cRelInst
                        lcSiglas = rsPers!cSiglas
                        
                        If lcTipPers <> "1" Then
                            lcNuDoTr = IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                            lcTiDoTr = IIf(IsNull(rsPers!NroRUC), "", 2)
                            lcTiDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                            lcNuDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                            lcCodUnico = lcNuDoTr 'Codigo Unico --Efrain
                        Else
                            lcNuDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                            lcTiDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", 2)
                            'lcTiDoCi = Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                            lcTiDoCi = rsPers!TipoDoc
                            lcNuDoCi = Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                            lcCodUnico = lcNuDoCi  'Codigo Unico --Efrain
                            lcCodRegPub = ""
                            'lcMagEmp = "0"
                            lcMagEmp = rsPers!cMagPersNat 'JACA 20110428
                            lcSiglas = ""
                        End If
                        
                        If Val(lcTipPers) >= 3 Then 'ARCV 17-06-2007
                            lcTipPers = "2"
                        End If
    
                        lcResid = "1"
    
                        If rsPers!cRelInst = "A" Then
                            lcAccionista = "1"
                        Else
                            lcAccionista = "0"
                        End If
    
                        If Not IsNull(rsPers!cRelInst) Then
                           Select Case Trim(rsPers!cRelInst)
                            Case "N", "O"  'ninguno/ no indicado
                                lcRelInst = "0"
                            Case "D"        'director
                                lcRelInst = "1"
                            Case "F"        'funcionario
                                lcRelInst = "2"
                            Case "T"        'trabajador
                                lcRelInst = "3"
                            Case "A"
                                lcRelInst = "0"
                            End Select
                        Else
                            lcRelInst = "0"
                        End If
    
                        lcPaisNac = "PE"
                        lcGenero = rsPers!Genero
                        '**DAOR 20071121
                        If rsPers!cTipPers >= 2 Then 'Para personas jurídicas Estado Civil=0
                            lcEstadoCiv = "0"
                        Else
                            lcEstadoCiv = "S"
                            Select Case rsPers!cEstadoCiv
                                Case Null: lcEstadoCiv = "0"
                                Case "1": lcEstadoCiv = "S" '1
                                Case "2": lcEstadoCiv = "C" '2
                                Case "3": lcEstadoCiv = "V" '4
                                Case "4": lcEstadoCiv = "D" '3
                                Case "5": lcEstadoCiv = "C" '2
                                Case "6": lcEstadoCiv = "S" '1
                            End Select
                            If lcEstadoCiv = "D" And Len(lsCas) > 0 Then
                                lcEstadoCiv = "C"
                            End If
                        End If
                    End If
                    rsPers.Close
                    Set rsPers = Nothing '***** LUCV20170411, Agregó
                End If
                
                '->** LUCV20180904, Comentó y Agregó. Según Recomendación SBS
    '            'JUEZ 20130809 ************************************************
    '            If rs!NroIFIS <= 3 And rs!nMinorista = 1 Then 'hasta en 3 ifis y que no sean mediana, grande y corporativa
    '                lcCondEndeuda = "0"
    '            ElseIf rs!NroIFIS > 3 And rs!nMinorista = 1 Then 'mas 3 ifis y que no sean mediana, grande y corporativa
    '                lcCondEndeuda = "1"
    '            ElseIf rs!nMinorista = 0 Then 'que sean mediana, grande y corporativa
    '                lcCondEndeuda = "9"
    '            End If
    '            'END JUEZ *****************************************************
                lcCondEndeuda = rs!lcCondEndeuda
                '<-** Fin LUCV20180904
                
                '******************** CALIFICACION DE LA PERSONA **************
                lcCalificacion = rs!cCalif
                '** Verifica calificacion
                If lcCalificacion = "" Then
                    ReDim Preserve psMensaje(i)
                    psMensaje(i - 1) = "Calificacion en blanco Encontrada en Cliente [" & rs!cPersCod & "] " & lsNombrePersona & Chr(13) & " por favor verifique los datos del cliente o consulte a sistemas"
                End If
                
                '***** Riesgo Cambiario
                lcIndRCambiario = IIf(IsNull(rs!cRCambiario), "0", rs!cRCambiario)
                '***** Indicador de Atraso
                'ARCV 16-06-2007 : Descomentar una vez enviada la Metodologia a la SBS
                'lcIndAtraso = IIf(IsNull(rs!cIndicadorAtrasoSBS), "X", rs!cIndicadorAtrasoSBS)
                lcIndAtraso = "X" 'Comentar una vez enviada la Metodologia a la SBS
                '-----------------
                ' By Capi 13112007
                'lcCalificacion = IIf(IsNull(rs!cCalif), "", rs!cCalif)
                If lcCalificacion = " " Then
                    lcCalificacion = "8"
                End If
                'End By
    
                '***** Calificac Interna
                'ARCV 16-06-2007 : Descomentar una vez enviada la Metodologia a la SBS
                lcIndClasifInterna = "XXXXX" 'Comentar una vez enviada la Metodologia a la SBS
                
                lsSQL = "INSERT INTO " & psServConsol & lsNomTabla & psFecha & "01  " 'JUEZ 20150310
                lsSQL = lsSQL & "(cTipoFor,cTipoInf,cNumSec, cCodSBS,cPersCod, cCodUnico, cActEcon, cCodRegPub, "
                lsSQL = lsSQL & " cTidoTr, cNudoTr, cTiDoci, cNuDoci, cTipPers, cResid, cCalifica, cMagEmp, "
                lsSQL = lsSQL & " cAccionista, cRelInst,cPaisNac, cSiglas,cPersNom,cPersNomCom,cPersGenero,cPersEstado, "
                lsSQL = lsSQL & " cApePat, cApeMat ,cApeCasada, cNombre1, cNombre2, "
                lsSQL = lsSQL & " cIndRCambiario, cIndAtrasoDeudor, cClasifInterna,cPersCodGrupEco,dFecNac,cTiDociComp,cNuDociComp,cCaliSinAlinea,cCondEndeuda) " 'JUEZ 20130809 Se agregó cCondEndeuda
                lsSQL = lsSQL & " VALUES('" & lsTipoDocRCD & "','" & lsTipoInfRCD & "',Null, "
                lsSQL = lsSQL & "'" & lcCodSBS & "','" & lsCodPers & "','" & lcCodUnico & "', "
                lsSQL = lsSQL & "'" & lcActEcon & "'" & "," & "'" & lcCodRegPub & "'" & ","
                lsSQL = lsSQL & IIf(lcTiDoTr = "Null", "Null", "'" & Trim(lcTiDoTr) & "'") & ","
                lsSQL = lsSQL & IIf(lcNuDoTr = "Null", "Null", "'" & Trim(lcNuDoTr) & "'") & ",'"
                lsSQL = lsSQL & Trim(IIf(IsNull(lcTiDoCi), "1", lcTiDoCi)) & "',"
                lsSQL = lsSQL & Trim(IIf(IsNull(lcNuDoCi), "NULL", "'" & Trim(lcNuDoCi) & "'")) & ",'"
                lsSQL = lsSQL & Trim(lcTipPers) & "','"
                lsSQL = lsSQL & lcResid & "','" & lcCalificacion & "','" & lcMagEmp & "','"
                lsSQL = lsSQL & lcAccionista & "','" & lcRelInst & "','" & lcPaisNac & "',"
                lsSQL = lsSQL & "'" & lcSiglas & "'" & ",'"
                lsSQL = lsSQL & Trim(lcNomPers) & "','" & Trim(lcNomPersComp) & "','" & lcGenero & "','" & lcEstadoCiv & "','"
                lsSQL = lsSQL & Trim(lsPat) & "','" & Mid(Trim(lsMat), 1, 39) & "','" & Mid(Trim(lsCas), 1, 39) & "','" & Mid(Trim(lsNom1), 1, 39) & "','" & Mid(Trim(lsNom2), 1, 39) & "','"
                lsSQL = lsSQL & lcIndRCambiario & "','" & lcIndAtraso & "','" & lcIndClasifInterna & "', '" & rs!nGrupoCod & "','" & Format(rs!dPersNacCreac, "YYYY/MM/DD") & "' ,'" & rs!cTpoDocCom & "','" & rs!cNumDocCom & "','" & rs!cCalCMAC & "','" & lcCondEndeuda & "') " 'JUEZ 20130809 Se agregó lcCondEndeuda
                oCon.Ejecutar (lsSQL)
            End If
                        
            '*******************************************
            '**  DATOS DE SALDOS DE CUENTAS - RCDvc02 **
            '*******************************************
            'ARCV 16-06-2007 (Agencia 08 como 09)
            'ALPA 20100627 B2***********************************************************************************************
            lsCredito = IIf(rs!cAgeCodB = "08", Left(rs!cCtaCod, 3) & "09" & Mid(rs!cCtaCod, 6, 13), rs!cCtaCod)
            '***************************************************************************************************************
            
            lcCalificacion = rs!cCalif
            lbRefinan = IIf((rs!nPrdEstado = gColocEstRefNorm Or rs!nPrdEstado = gColocEstRefMor Or rs!nPrdEstado = gColocEstRefVenc), True, False)
            
            If rs!nPrdEstado = gColocEstRecVigJud Or rs!nPrdEstado = gColocEstRecVigCast _
               Or rs!nPrdEstado = 2205 Or rs!nPrdEstado = 2206 Then
                lbCobJud = True
                lbCastigado = IIf(rs!nPrdEstado = gColocEstRecVigCast Or rs!nPrdEstado = 2206, True, False)
                lbDemanda = IIf(rs!nDemanda = 1, True, False)
                lnSaldoCap = rs!nSaldoCap
                lnDiasAtraso = rs!nDiasAtraso
            Else
                lbCobJud = False
                lbCastigado = False
                lbDemanda = False
                lnSaldoCap = rs!nSaldoCap
                lnDiasAtraso = rs!nDiasAtraso
            End If
            
            'JUEZ 20141111 **********************************
            If rs!nPrdEstado = 2205 And rs!cSitCtb = 5 Then
                lbRefinan = True
                lbCobJud = False
            End If
            'END JUEZ ***************************************
            
            lsEstadoCredito = DevEstadoCredito(Mid(Trim(rs!cTpoCredCod), 1, 1), lnDiasAtraso, lbCobJud)
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000138", Mid(rs!cCtaCod, 6, 3)) Then
            'If Mid(rs!cCtaCod, 6, 3) = "515" Then
            '**ARLO20180712 ERS042 - 2018
                If lnDiasAtraso <= 30 Then
                    lsEstadoCredito = "VIG"
                Else
                    lsEstadoCredito = "VEN"
                End If
            End If
            If lnDiasAtraso < 0 Then lnDiasAtraso = 0
            
            '***************************************************************
            ' Genero cuentas Contables de los Creditos que tenga el Cliente
            '****************************************************************
            ' ******** Cuenta de Saldos de Capital
            lsOpeEnGarant = "02"
            lsCuentaCnt = ""
            
            'No afectar la Carta Fianza a las cuentas de la 14
            If Mid(rs!cTpoCredCod, 2, 2) = "81" Then
                'By Capi 08112007 para que genere la cuenta 27
                 If lcCalificacion > "0" Then
                    lsCuentaCnt = "27" & Mid(lsCredito, 9, 1) & "101" & rs!cTpoCredCodCnt
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, rs!nProvision, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                 End If
                'End By
                lsCuentaCnt = "71" & Mid(lsCredito, 9, 1) & "2" '& "06"
                Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
            Else
                '**ARLO20180712 ERS042 - 2018
                Set objProducto = New COMDCredito.DCOMCredito
                If (((Mid(rs!cTpoCredCod, 1, 1) = "6" Or Mid(rs!cTpoCredCod, 1, 1) = "7") And rs!cTpoCredCod <> "755") Or objProducto.GetResultadoCondicionCatalogo("N0000139", Mid(rs!cCtaCod, 6, 3))) And lsEstadoCredito = "VEN" Then 'ARLO20190213
                'If (((Mid(rs!cTpoCredCod, 1, 1) = "6" Or Mid(rs!cTpoCredCod, 1, 1) = "7") And rs!cTpoCredCod <> "755") Or (Mid(rs!cCtaCod, 6, 3) = "515")) And lsEstadoCredito = "VEN" Then
                '**ARLO20180712 ERS042 - 2018
                    'Inserta la parte VENCIDA
                    If lnDiasAtraso <= 90 Then 'Vence el Capital de Cuota
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nCapVenc, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                        'Inserta la parte VIGENTE
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, "VIG", lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap - rs!nCapVenc, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    Else
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    End If
                ElseIf Mid(rs!cTpoCredCod, 1, 1) = "8" And lsEstadoCredito = "VEN" Then
                    'Inserta la parte VENCIDA
                    If lnDiasAtraso <= 90 Then ' Vence el Capital de Cuota
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nCapVenc, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                        ' Inserta la parte VIGENTE
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, "VIG", lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        '************************************************************
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap - rs!nCapVenc, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    Else
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    End If
                Else 'Inserta Capital
                    'By capi 13112007 para que no afecte a la cuenta 14
                    If rs!nPrdEstado <> "2108" And rs!nPrdEstado <> "2060" Then
                        lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, lbCobJud, lbDemanda, lbCastigado, rs!nRealizoPagoCastigado, rs!cTpoCredCod, rs!cTpoCredCodCnt, rs!cSC, rs!cCD, rs!cIK, rs!cSitCtb, rs!cIFTpo)
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    End If
                End If
                
                'Cuenta de Intereses *******************
                lsOpeEnGarant = "02"
                lnIntDeveng = IIf(IsNull(rs!nIntDeveng), 0, rs!nIntDeveng)
                lnIntDevengCC = IIf(IsNull(rs!nIntDevengCC), 0, rs!nIntDevengCC)
                lnSaldoGasto = IIf(IsNull(rs!nSaldoGasto), 0, rs!nSaldoGasto) 'JUEZ 20121002
                lnMontoReprogDesNat = IIf(IsNull(rs!nMontoReprogxDesNat), 0, rs!nMontoReprogxDesNat) 'EJVG20121212
                lnMontoReprogZonaAfectInunda = IIf(IsNull(rs!nMontoReprogxZonaAfectInunda), 0, rs!nMontoReprogxZonaAfectInunda) 'LUCV20170510, Agrego por Operacion: 100925
                nSaldoIntSuspensoDeven = 0
                If rs!cSitCtb = "1" And (rs!cCalif = 0 Or rs!cCalif = 1 Or rs!cCalif = 2) And lbRefinan = False Then
                    nLogDeven = 1
                Else
                    nLogDeven = 0
                End If
                
                If lsEstadoCredito = "VIG" And lbRefinan = False And rs!cSitCtb = "1" And nLogDeven = 1 Then
                    lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "8" & rs!cTpoCredCodCnt '& rs!cIFTpo
                    nSaldoIntSuspensoDeven = lnIntDeveng + rs!nIntSusp 'CC
                 ElseIf lbRefinan = False And (rs!cSitCtb = "5" Or rs!cSitCtb = "2") And nLogDeven = 0 Then
                    lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "402"
                    nSaldoIntSuspensoDeven = rs!nIntSusp + rs!nIntDeveng
                 ElseIf (lbRefinan = True And rs!cSitCtb = "4" And nLogDeven = 0) Or (lbRefinan = False And rs!cSitCtb = "1" And nLogDeven = 0) Then
                    lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "401"
                    nSaldoIntSuspensoDeven = rs!nIntSusp + rs!nIntDeveng
                'Inicio cambio intereses suspenso
                 ElseIf lbCobJud = True And nLogDeven = 0 Then
                    If lbCastigado Then
                        If rs!nRealizoPagoCastigado Then
                            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "925"
                            nSaldoIntSuspensoDeven = lnIntDeveng 'CC
                                
                            If lnSaldoGasto > 0 Then
                                Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "925", 0, lnSaldoGasto, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                            End If
                        Else
                            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "304"
                            nSaldoIntSuspensoDeven = lnIntDeveng 'CC
                                
                            If lnSaldoGasto > 0 Then
                                Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "305", 0, lnSaldoGasto, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                            End If
                        End If
                    Else
                        If lbDemanda Then
                            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "403"
                            nSaldoIntSuspensoDeven = rs!nIntSusp + rs!nIntDeveng
                        ElseIf (rs!cSitCtb = "5" Or rs!cSitCtb = "2") And nLogDeven = 0 Then
                            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "402"
                            nSaldoIntSuspensoDeven = rs!nIntSusp + rs!nIntDeveng
                        End If
                    End If
                Else
                    If nLogDeven = 0 And (rs!cSitCtb = "5" Or rs!cSitCtb = "2") Then
                        lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "402"
                        nSaldoIntSuspensoDeven = rs!nIntSusp + rs!nIntDeveng
                    End If
                End If
                'fin cambio intereses suspenso
                
                If nSaldoIntSuspensoDeven > 0 Then
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, nSaldoIntSuspensoDeven, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
        
                'Interes Adelantados Pignoraticio
                'Provision Especifica
                If Val(lcCalificacion) > 0 And lbCastigado = False Then
                    lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "9" & rs!cTpoCredCodCnt & "01"
                    lnProvision = Format(IIf(IsNull(rs!nProvision), 0, rs!nProvision), "#.00")
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnProvision, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                'Provision por SobreEndeudamiento
                If rs!nProvisionSobreendeu > 0 Then
                    lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "9" & rs!cTpoCredCodCnt & "06"
                    lnProvision = Format(IIf(IsNull(rs!nProvisionSobreendeu), 0, rs!nProvisionSobreendeu), "#.00")
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnProvision, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                
                'Provision Especifica
                If rs!nProvisionRCC > 0 Then
                    lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "9" & rs!cTpoCredCodCnt & "05"
                    lnProvision = Format(IIf(IsNull(rs!nProvisionRCC), 0, rs!nProvisionRCC), "#.00")
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnProvision, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                
                'EJVG20121212 *** Reprog x Desastre Nat o Conflicto Social
                If lnMontoReprogDesNat > 0 Then
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "927", 0, lnMontoReprogDesNat, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                'END EJVG *******
                
                'LUCV20170510 *** Reprogramación Por Zona Afectada Por Inundación
                If lnMontoReprogZonaAfectInunda > 0 Then
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "936", 0, lnMontoReprogZonaAfectInunda, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                'Fin LUCV20170510 *******
                
                'JUEZ 20130809 Creditos Hipotecarios otorgados a partir del 01/01/2013 **********
                If rs!nHipotecario810928 = 1 Then
                    If rs!nDatoVivienda = 1 Then 'si es para primera vivienda "81092801"
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "92801", 0, lnSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    Else 'caso contrario "81092803"
                        Call GrabaEnRCDSaldos(psFecha, lsCodPers, "81" & Mid(lsCredito, 9, 1) & "92803", 0, lnSaldoCap, rs!cTpoCreditoEquivalente, pnTipoCambio, "06", lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                    End If
                End If
                'END JUEZ ***********************************************************************
            End If
        
             If rs!cTpoCredCod = "755" Or rs!nPrdEstado = 2060 Then 'Pignoraticios y By Capi 07112007 los canceladosxrefinanciacion cta 29010104
                lsCondEspecial = "06"
                lnDiasGarantia = 0
                lnMontoGar2 = lnSaldoCap
                '**By Capi 06112007, para que incluya la cuenta 16020102 y , las demas subcuentas
                '**de la 16 no se considero porque el control no esta contemplado en el
                '**sistema.
                If rs!nPrdEstado = 2108 Then 'Adjudicados
                    lsCuentaCnt = "16120102" '& Mid(lsCredito, 4, 2)
                ElseIf rs!nPrdEstado = 2060 Then
                    lsCuentaCnt = "29" & Mid(lsCredito, 9, 1) & "10104" '& Mid(lsCredito, 4, 2)
                End If
            End If
            '**************************************************************************************
                
            
            lsSQL = "select rtrim(cTipoGarant) as cTipoGarant,nMoneda, nMonGarantia,cCtaContCod from " & psServConsol & "#TmpGarantiasRCD where cCtaCod='" & lsCredito & "'" '***** LUCV20170411, Comentó
            'lsSQL = "exec stp_sel_RecuperaDatosGarantiasRCD '" & Format(pdFecDataFM, "yyyymmd") & "', '" & lsCredito & "'" '***** LUCV20170411, Agregó
            Set rsSaldosCont = oCon.CargaRecordSet(lsSQL)
    
            Do While Not rsSaldosCont.EOF
                lsCuentaCnt = ""
                lsCondEspecial = "06"
                lnDiasGarantia = 0
    
                lsCuentaCnt = rsSaldosCont!cCtaContCod
                lnMontoGar2 = rsSaldosCont!nMonGarantia
                If rsSaldosCont!cTipoGarant = "1" Then
                    lsCondEspecial = "04"
                End If
   
                If lsCuentaCnt <> "" Then
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasGarantia, lnMontoGar2, rs!cTpoCreditoEquivalente, pnTipoCambio, lsCondEspecial, lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
                End If
                rsSaldosCont.MoveNext
                DoEvents
            Loop
            rsSaldosCont.Close
            Set rsSaldosCont = Nothing '***** LUCV20170411, Agregó
            'End If
            '***********************************************************************************
            lsSQL = "select rtrim(cTipoGarant) as cTipoGarant,nMoneda, nMonGarantia,cCtaContCod from " & psServConsol & "#TmpGarantiasRCDCF where cCtaCod='" & lsCredito & "'" '***** LUCV20170411, Comentó
            'lsSQL = "exec stp_sel_RecuperaDatosGarantiasRCDF '" & Format(pdFecDataFM, "yyyymmd") & "', '" & lsCredito & "'" '***** LUCV20170411, Agregó
            Set rsSaldosCont = oCon.CargaRecordSet(lsSQL)
            Do While Not rsSaldosCont.EOF
               lsCuentaCnt = ""
               lsCondEspecial = "06"
               lnDiasGarantia = 0
               'ALPA20160510********************************
               lsCuentaCnt = rsSaldosCont!cCtaContCod '"84" & rsSaldosCont!nMoneda & "409"
               lnMontoGar2 = rsSaldosCont!nMonGarantia
               '********************************************
               If lsCuentaCnt <> "" Then
                    Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, lnDiasGarantia, lnMontoGar2, rs!cTpoCreditoEquivalente, pnTipoCambio, lsCondEspecial, lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, rs!cCtaCod, pbRCDTransf)
               End If
               rsSaldosCont.MoveNext
               DoEvents
            Loop
            rsSaldosCont.Close '***** LUCV20170411, Agregó
            Set rsSaldosCont = Nothing '***** LUCV20170411, Agregó
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing '***** LUCV20170411, Agregó
    
    If Not pbRCDTransf Then 'JUEZ 20150310
        'JUEZ 20121001 *************************************************************************************************
        Dim rsDif As ADODB.Recordset
        lsSQL = "exec stp_sel_RecuperaDatosIntDiferidosRCD '" & Format(pdFecDataFM, "yyyymmdd") & "'"
        
        Set rsDif = oCon.CargaRecordSet(lsSQL)
        Do While Not rsDif.EOF
            Call GrabaEnRCDSaldos(psFecha, rsDif!cPersCod, rsDif!CtaCont, rsDif!nDiasAtraso, rsDif!nSaldo, rsDif!cTpoCreditoEquivalente, pnTipoCambio, rsDif!cCondEspCta, rsDif!cCondDisponib, psServConsol, rsDif!cAgeCod, rsDif!cUbiGeoSBS, rsDif!cCtaCod, pbRCDTransf)
           rsDif.MoveNext
        DoEvents
        Loop
        rsDif.Close
        'END JUEZ ******************************************************************************************************
    End If

    'Actualiza el Total en RCDvcXXXXX02 a partir de RCDvcXXXXX02T
    'JUEZ 20150310 Se modificó "RCDvc" por lsNomTabla
    lsSQL = "INSERT INTO " & psServConsol & lsNomTabla & psFecha & "02 (cTipoFor,cTipoInf,cNumSec, cCodAge," _
          & " cUbicGeo, cCtaCnt, cTipoCred,nSaldo, nCondDias, cCondEspCta, cCondDisponib, cPersCod, cCtaCod) " _
          & " SELECT cTipoFor, cTipoInf, null, cCodAge, cUbicgeo, cCtaCnt, ctipoCred, " _
          & " sum(nsaldo), nCondDias , cCondEspCta, cCondDisponib, cPersCod, cCtaCod " _
          & " From " & psServConsol & lsNomTabla & psFecha & "02T " _
          & " Group By cTipoFor, cTipoInf, cCodAge, cUbicgeo, cCtaCnt, cTipoCred, nCondDias, cCondEspCta, cCondDisponib, cPersCod, cCtaCod  "
    oCon.Ejecutar lsSQL
    
    ' Actualiza el Total en RCDvcXXXXX03
    'JUEZ 20150310 Se modificó "RCDvc" por lsNomTabla
    lsSQL = "INSERT INTO " & psServConsol & lsNomTabla & psFecha & "03 (cTipoFor,cTipoInf,cNumSec,cCtaCnt, cTipoCred,nSaldo, " _
          & " nCondDias, cCondEspCta, cCondDisponib) " _
          & " Select '2','2',Null, cCtaCnt, cTipoCred, Sum(nSaldo), nCondDias, cCondEspCta, max(cConddisponib) " _
          & " From " & psServConsol & lsNomTabla & psFecha & "02 " _
          & " Group By cCtaCnt, cTipoCred, nCondDias, cCondEspCta " _
          & " Order By cCtaCnt, cTipocred, nconddias, cCondEspCta "
    oCon.Ejecutar lsSQL

Set rs = Nothing
End Sub

Private Function VerApostrofe(ByVal psCadena As String) As String
    VerApostrofe = Replace(psCadena, "'", "''", , , vbTextCompare)
End Function

Private Function fReemplazaCaracterEspecial(ByVal psNombre As String) As String
Dim lsNomMod As String

lsNomMod = psNombre
If InStr(1, lsNomMod, "'", vbTextCompare) <> 0 Then
    lsNomMod = Replace(lsNomMod, "'", "''", , , vbTextCompare)
End If
lsNomMod = Replace(lsNomMod, "Ñ", "#", , , vbTextCompare)
lsNomMod = Replace(lsNomMod, "ñ", "#", , , vbTextCompare)
lsNomMod = Replace(lsNomMod, "-", "", , , vbTextCompare)
lsNomMod = Replace(lsNomMod, "|", "", , , vbTextCompare)
lsNomMod = Replace(lsNomMod, ".", " ", , , vbTextCompare)
    'al final
lsNomMod = Replace(lsNomMod, "   ", " ", , , vbTextCompare)
lsNomMod = Replace(lsNomMod, "  ", " ", , , vbTextCompare)
lsNomMod = Trim(lsNomMod)

fReemplazaCaracterEspecial = lsNomMod
End Function

'**Modificado por DAOR 20080811, se aumentó el parametro pbRealizoPagoCastigado
'Function GeneraCuentaCapital(ByVal pCredito As String, ByVal pEstadoCredito As String, ByVal pRefinan As Boolean, ByVal pCobJud As Boolean, ByVal pDemanda As Boolean, ByVal pCastigado As Boolean, Optional ByVal pbRealizoPagoCastigado As Boolean = False) As String
Function GeneraCuentaCapital(ByVal pCredito As String, ByVal pEstadoCredito As String, ByVal pRefinan As Boolean, ByVal pCobJud As Boolean, ByVal pDemanda As Boolean, ByVal pCastigado As Boolean, Optional ByVal pbRealizoPagoCastigado As Boolean = False, Optional ByVal psTpoCredCod As String, Optional ByVal psTpoCredCodCnt As String = "", Optional ByVal psSC As String = "", Optional ByVal psCD As String = "", Optional ByVal psIK As String = "", Optional ByVal sSitCtb As String, Optional ByVal psIFTpo As String = "") As String
Dim lsCuenta As String

Dim lsSituacion As String
Dim lsCondicion As String
Dim lsIk As String
Dim lsCuota As String

lsCondicion = psCD
lsIk = psIK
If pCobJud = True Then    ' Cobranza Judicial
    
    If pCastigado = True Then ' Si castigado
        '**Modificado por DAOR 20080811 ************************
        If pbRealizoPagoCastigado Then
            lsCuenta = "81" & Mid(pCredito, 9, 1) & "925"
        Else
            lsCuenta = "81" & Mid(pCredito, 9, 1) & "302"
        End If
        '*******************************************************
        
        GeneraCuentaCapital = lsCuenta
        Exit Function
    Else
        If pDemanda = True Then
            lsSituacion = "6"
            'lsCondicion = IIf(pRefinan = False , "06",  IIf( Mid(pCredito, 3, 1) = "3", "19", "1929"))
            'ALPA 20100627 B2******************************************
            'lsCondicion = IIf(Mid(pCredito, 3, 3) = "423", "23", "06")
            'lsCondicion = IIf(psTpoCredCod = "853", "23", "06")--Borrar despues
            '**********************************************************
            
            'By capi 13032009 se agrego para administrativos
            'ALPA 20100627 B2****************************************************
            'lsCondicion = IIf(Mid(pCredito, 6, 3) = "320", "2009", lsCondicion)
            'lsCondicion = IIf(psTpoCredCod = "756", "2009", lsCondicion)--Borrar despues
            
            If pRefinan = False Then
                lsCondicion = psCD
            Else
                If pEstadoCredito = "VEN" Then
                    lsCondicion = "19"
                Else
                    lsCondicion = psCD
                End If
            End If
            '********************************************************************
        Else
            'By Capi 11022009 para controlar creditos administrativos
            'lsSituacion = "5"
            
            
            'lsSituacion = IIf(Mid(pCredito, 6, 3) = "320", "6", "5")
            lsSituacion = IIf(psTpoCredCod = "756", "6", "5")
            'End by
            'lsCondicion = IIf(pRefinan = False, "06", IIf(Mid(pCredito, 3, 1) = "3", "19", "1929")) ' CAMBIO SILVITA
            '**Comentado por DAOR 20070912
            'lsCondicion = IIf(Mid(pCredito, 3, 3) = "423", "23", "06")
            
            '**DAOR 20070912 **********************************************
            'lsCondicion = IIf(Mid(pCredito, 6, 3) = "320", "2009", "06")
            '**************************************************************
            
        End If
    End If
    
'ElseIf Mid(pCredito, 6, 3) = "305" Then  ' Prendario
ElseIf psTpoCredCod = "755" Then  ' Prendario
    If pEstadoCredito = "VIG" Then
        lsSituacion = "1"
        'lsCondicion = "13"
    Else
        lsSituacion = "5"
        'lsCondicion = "13"
    End If
'ElseIf Mid(pCredito, 6, 3) = "320" Then ' Administrativos
ElseIf psTpoCredCod = "756" Then ' Administrativos
    If pEstadoCredito = "VIG" Then
        lsSituacion = "1"
        'lsCondicion = "2009"
    Else
        lsSituacion = "5"
        'lsCondicion = "2009"
    End If

'ElseIf Mid(pCredito, 6, 3) = "423" Then ' Mivivienda
ElseIf psTpoCredCod = "853" Then ' Mivivienda
    If pEstadoCredito = "VIG" Then
        lsSituacion = "1"
        'lsCondicion = "23"
    Else
        lsSituacion = "5"
        'lsCondicion = "23"
    End If
    
Else   ' Comerc - Pyme - Consumo - HipoteCaja
    If pEstadoCredito = "VIG" Then
        If pRefinan = False Then
            lsSituacion = "1"
            
            'lsCondicion = "06"
            'lsCondicion = "0601"
        Else  ' Refinanciado
            lsSituacion = "4"
            'lsCondicion = "06"
        End If
    Else    ' Vencido
        If pRefinan = False Then
            lsSituacion = "5"
            'lsCondicion = "06"
        Else  ' Refinanciado
            'If Mid(pCredito, 6, 1) = "3" Or Mid(pCredito, 6, 1) = "4" Then ' Para Consumo no ha cambiado los refinanciados
            If Mid(psTpoCredCod, 1, 1) = "6" Or Mid(psTpoCredCod, 1, 1) = "7" Or Mid(psTpoCredCod, 1, 1) = "8" Then   ' Para Consumo no ha cambiado los refinanciados
                lsSituacion = "5"
             '   lsCondicion = "1906"
            Else
                lsSituacion = "5"
                'lsCondicion = "1929"  ' Ojo
              '  lsCondicion = "1906"  ' Ojo
            End If
        End If
    End If
End If
    If pRefinan = False Then
        lsCondicion = psCD
        lsIk = psIK
    Else
        If pEstadoCredito = "VEN" Then
            lsCondicion = "19"
            lsIk = "06"
        Else
            lsCondicion = psCD
            lsIk = psIK
        End If
    End If
'***  [14][M][Sit][TC][Cond]
'lsCuenta = "14" & Mid(pCredito, 9, 1) & lsSituacion & "0" & Mid(pCredito, 6, 1) & lsCondicion
lsCuenta = "14" & Mid(pCredito, 9, 1) & lsSituacion & psTpoCredCodCnt & lsCondicion & IIf(psTpoCredCod = "755", "00", lsIk) '& psIFTpo

'*** Agregar el Tipo de Cuota 2006/12/15 (2007-02-08 layg)
'ALPA 20100709 B2***********************************************************
'If pRefinan = False Or (pRefinan = True And lsSituacion = "4") Then
'    'lsCuenta = lsCuenta & fVarCU(Mid(pCredito, 6, 3))
'    lsCuenta = lsCuenta & fVarCU(psTpoCredCod)
'End If
'***************************************************************************

'By capi 11022009 para pasar saldos de 14010406 1401040601
If lsCuenta = "14110406" Or lsCuenta = "14210406" Then
    lsCuenta = Left(lsCuenta, 2) & Mid(lsCuenta, 3, 1) & Right(lsCuenta, 5) & "01"
End If
'End by

'JUEZ 20121224 **************************************************************
If psTpoCredCod = "853" Or psTpoCredCod = "154" Or psTpoCredCod = "254" Or _
   psTpoCredCod = "354" Or psTpoCredCod = "454" Or psTpoCredCod = "554" Or _
   psTpoCredCod = "854" Then 'JUEZ 20130304 Se agregó 854: Techo Propio
   lsCuenta = Left(lsCuenta, 8)
End If
'END JUEZ *******************************************************************

GeneraCuentaCapital = lsCuenta

End Function

Private Function fVarCU(ByVal pTipoCredito As String) As String
Dim lsVar As String
'If pTipoCredito = "320" Then
'    lsVar = ""
'ElseIf pTipoCredito = "305" Then
'    lsVar = ""
'ElseIf pTipoCredito = "423" Then
'    lsVar = ""
'ElseIf pTipoCredito = "401" Or pTipoCredito = "402" Then
'    lsVar = ""
'ElseIf Mid(pTipoCredito, 1, 1) = "3" Then
'    lsVar = "09"
'Else
'    lsVar = "02"
'End If
If pTipoCredito = "756" Then
    lsVar = ""
ElseIf pTipoCredito = "755" Then
    lsVar = ""
ElseIf pTipoCredito = "853" Then
    lsVar = ""
ElseIf pTipoCredito = "851" Then
    lsVar = ""
ElseIf Mid(pTipoCredito, 1, 1) = "6" Or Mid(pTipoCredito, 1, 1) = "7" Then
    lsVar = "09"
Else
    lsVar = "02"
End If
fVarCU = lsVar
End Function

Function DevEstadoCredito(ByVal pCredito As String, ByVal pDiasAtraso As Integer, ByVal pCobJud As String) As String
Dim lsEstado As String

If pCobJud = True Then
    lsEstado = "COJ"
Else
'    Select Case Mid(pCredito, 6, 1)
'            Case "1" ' Comercial
'                If pDiasAtraso <= 15 Then
'                    lsEstado = "VIG"
'                Else
'                    lsEstado = "VEN"
'                End If
'            Case "2", "3", "4"   'Mes, Consumo, Hipotecario
'                If pDiasAtraso <= 30 Then
'                    lsEstado = "VIG"
'                Else
'                    lsEstado = "VEN"
'                End If
'    End Select
 Select Case pCredito
            Case "1", "2", "3" ' Comercial
                If pDiasAtraso <= 15 Then
                    lsEstado = "VIG"
                Else
                    lsEstado = "VEN"
                End If
            Case "4", "5", "6", "7", "8" 'Mes, Consumo, Hipotecario
                If pDiasAtraso <= 30 Then
                    lsEstado = "VIG"
                Else
                    lsEstado = "VEN"
                End If
    End Select
End If
DevEstadoCredito = lsEstado
End Function

'***********************************
'**  Graba en RCD02T ********
'***********************************
Sub GrabaEnRCDSaldos(ByVal pFecha As String, ByVal pcCodPers As String, _
                    ByVal pcCodCnt As String, ByVal pnDiasAtraso As Integer, ByVal pnSaldo As Currency, _
                    ByVal pTipoCredito As String, ByVal pnTipCambio As Double, _
                    ByVal pcCondEspecial As String, ByVal pcCondDisponib As String, _
                    ByVal psServConsol As String, _
                    ByVal psAgencia As String, ByVal psUbicaGeoRCD As String, Optional ByVal psCtaCod As String = "", _
                    Optional ByVal pbRCDTrans As Boolean = False)
                    'JUEZ 20150310 Se agregó pbRCDTransf

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim lbExiste As Boolean
Dim lnSaldoSoles As Currency
Dim lnAgencia As Integer
Dim rsVerif As New ADODB.Recordset ' Para Verificar

    If pnSaldo <= 0 Then
        Exit Sub
    End If

'Agencia
lnAgencia = Val(psAgencia)

'****  Cambio a Moneda Soles  ****
'->***** LUCV20180904 Agregó condicional, Según Provisiones por SobreEnd.
If Mid(pcCodCnt, 1, 2) = "14" And (Mid(pcCodCnt, 3, 2) = "19" Or Mid(pcCodCnt, 3, 2) = "29") And Mid(pcCodCnt, 7, 2) = "06" And Len(pcCodCnt) = 8 Then
    lnSaldoSoles = Format(pnSaldo, "#0.00")
Else
    If Mid(pcCodCnt, 3, 1) = "1" Then
        lnSaldoSoles = Format(pnSaldo, "#0.00")
    Else
        lnSaldoSoles = Format(pnSaldo * pnTipCambio, "#0.00")
    End If
End If
'<-***** Fin LUCV20180901

'*******************************************
'**  SALDOS POR CUENTA CONTABLE  - RCDvc02 **
'*******************************************
' -----** Graba en RCDvc02
'lsSQL = "SELECT COUNT(cPersCod) hay From " & psServConsol & "RCDvc" & pFecha & "02  WHERE cPersCod ='" & pcCodPers & "' " _
      & "AND cCtaCnt ='" & pcCodCnt & "' AND nCondDias = " & pnDiasAtraso & " AND cTipoCred = '" & pTipoCredito & "' "
'Set rs = oCon.CargaRecordSet(lsSQL)

'lbExiste = IIf(rs!hay > 0, True, False)
'rs.Close
'Set rs = Nothing

'If lbExiste = False Then
    'JUEZ 20150310 Se reemplazó "RCDvc" por pbRCDTrans
     lsSQL = "INSERT INTO " & psServConsol & IIf(pbRCDTrans, "RCDTvc", "RCDvc") & pFecha & "02T (cTipoFor,cTipoInf,cNumSec, " _
      & "cCodAge,cUbicGeo, cCtaCnt, cTipoCred,nSaldo, nCondDias,cCondEspCta, cCondDisponib, cPersCod, cCtaCod ) " _
      & " VALUES ('1','2',Null,'" & psAgencia & "','" & psUbicaGeoRCD & "','" _
      & pcCodCnt & "','" & pTipoCredito & "'," & lnSaldoSoles & "," & pnDiasAtraso & ",'" _
      & pcCondEspecial & "','" & pcCondDisponib & "','" & pcCodPers & "','" & psCtaCod & "') "
'Else
'     lsSQL = "UPDATE " & psServConsol & "RCDvc" & pFecha & "02  SET nSaldo = nSaldo + " & lnSaldoSoles & "" _
      & "WHERE cPersCod ='" & pcCodPers & "' AND  cCtaCnt = '" & pcCodCnt & "' " _
      & "AND nCondDias =" & pnDiasAtraso & " AND cTipoCred = '" & pTipoCredito & "' " _
      & "AND cCodAge ='" & psAgencia & "' " _
      & "AND cUbicGeo ='" & psUbicaGeoRCD & "' "
'End If
oCon.Ejecutar lsSQL

End Sub

Private Function EmiteCodigoAux(ByVal lsCodPers As String, ByVal psServConsol As String) As String

Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = "Select * from " & psServConsol & "RCDCodigoAux where cPersCod ='" & Trim(lsCodPers) & "'"
Set rs = oCon.CargaRecordSet(Sql)

If Not rs.EOF Then
    EmiteCodigoAux = Trim(rs!cCodAux)
Else
    EmiteCodigoAux = ""
End If
rs.Close
Set rs = Nothing
End Function

Sub CargaGarantiasTotal(ByVal psServConsol As String, _
                        ByRef prsGarant As ADODB.Recordset)
Dim Sql As String

Sql = "         SELECT  G.cPersCod, G.nTipoGarant, G.nMoneda,"
Sql = Sql & "           T.ccodcnt, T.ctippref, sum(G.nMontoRealiz) as nMonto"
Sql = Sql & "           FROM    " & psServConsol & "GarantiasConsol G"
Sql = Sql & "           LEFT JOIN " & psServConsol & "TIPOGARAN T ON T.nTpoGarantia = G.nTipoGarant"
Sql = Sql & "           WHERE   Exists (    select  GC.cPersCod"
Sql = Sql & "                               from    " & psServConsol & "GarantCredConsol GC"
Sql = Sql & "                                       JOIN " & psServConsol & "CreditoConsol C ON C.cCtaCod = GC.cCtaCod"
Sql = Sql & "                               where   C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206,2101,2104,2106,2107)"
Sql = Sql & "                                       and GC.cPersCod = G.cPersCod AND GC.cNumGarant = G.cNumGarant )"
Sql = Sql & "           group by G.cPersCod, nTipoGarant,G.nMoneda,  ccodcnt , ctippref"

Set prsGarant = oCon.CargaRecordSet(Sql)
prsGarant.Sort = "cPersCod"

End Sub

Private Function GeneraContableGarantiaNEW(ByVal psPersGarantCod As String, ByVal psFecha As String, _
                                           ByVal psProducto As String, ByVal prsGarant As ADODB.Recordset, _
                                           ByVal pnTipoCambio As Double, ByVal psServConsol As String, _
                                           ByVal psCodOfiInf As String, ByVal psUbicaGeoRCD As String, _
                                           ByVal psCtaNoPref As String, ByRef psMensaje As Variant)
Dim lsCtaCont As String
Dim lsDH As String
Dim lsTipoPref As String
Dim lsCodGaranCont As String
Dim lsCodAge As String
Dim lsTipoCred As String
Dim lsMoneda As String
Dim lnDebe As Currency
Dim lnHaber As Currency
Dim lsOpeCod As String
Dim lxCta As String
Dim lnMontoGar As Currency
Dim i As Integer
    
If prsGarant Is Nothing Then Exit Function

prsGarant.MoveFirst
prsGarant.Find "cPersCod ='" & psPersGarantCod & "'"
If Not prsGarant.EOF And Not prsGarant.BOF Then
    i = 0
    Do While Not prsGarant.EOF And prsGarant!cPersCod = psPersGarantCod
        lsCodAge = "01" 'Trim(prsGarant!cCodAge)
        lsCtaCont = psCtaNoPref
        lsMoneda = Trim(prsGarant!nMoneda)
        lsTipoPref = "02" 'Right(prsGarant!CPREF, 2)
        
        If Not IsNull(prsGarant!ccodcnt) Then
            lsCodGaranCont = prsGarant!ccodcnt
        Else
            i = i + 1
            ReDim Preserve psMensaje(i)
            psMensaje(i - 1) = "No se ha establecido cuenta Contable para tipo de Garantía tipo [" & prsGarant!nTipoGarant & "]"
        End If
        
        'lsTipoCred = prsGarant!nProducto
        lsTipoCred = psProducto
        Select Case lsTipoCred
            Case 2 'Microempresa
                lsTipoCred = "02"
            Case 3 'Consumo
                lsTipoCred = "03"
            Case 1 'Comerciales
                lsTipoCred = "01"
            Case 4 'Hipotecario
                lsTipoCred = "04"
        End Select
        
        lsCtaCont = Replace(lsCtaCont, "M", lsMoneda)
        lsCtaCont = Replace(lsCtaCont, "PF", lsTipoPref)
        lsCtaCont = Replace(lsCtaCont, "TG", lsCodGaranCont)
        lsCtaCont = Replace(lsCtaCont, "TC", lsTipoCred)
        lsCtaCont = Replace(lsCtaCont, "AG", lsCodAge)
        
        lxCta = Left(Trim(lsCtaCont), 8)
        If Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840409" Or Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840403" Then
            lxCta = Left(Trim(lsCtaCont), 6)
        End If
        lxCta = lxCta + String(14 - Len(lxCta), "0")
       
        lnMontoGar = prsGarant!nMonto
        If lnMontoGar > 0 Then
            'ARCV 17-12-2006
            'Call GrabaEnRCDSaldos(psFecha, psPersGarantCod, lxCta, 0, lnMontoGar, psProducto, pnTipoCambio, psServConsol, psCodOfiInf, psUbicaGeoRCD)
            Call GrabaEnRCDSaldos(psFecha, psPersGarantCod, lxCta, 0, lnMontoGar, psProducto, pnTipoCambio, "", "", psServConsol, psCodOfiInf, psUbicaGeoRCD)
            '------------
        End If
        
        prsGarant.MoveNext
        If prsGarant.EOF Then Exit Do
    Loop
Else
End If

End Function

Sub CambiaNombreRCD(ByVal psNomCli As String, _
                    ByRef psPat As String, _
                    ByRef psMat As String, _
                    ByRef psCas As String, _
                    ByRef psNom1 As String, _
                    ByRef psNom2 As String)

Dim lsNombre As String
Dim lnPos As String
Dim lsRazon As String

lsNombre = Trim(psNomCli)    ' && ELIMINAMOS BLANCOS A LA IZQUIERDA
lnPos = 0
psPat = ""
psMat = ""
psCas = ""
psNom1 = ""
psNom2 = ""

lsRazon = ""
lnPos = InStr(1, lsNombre, "/")
If lnPos > 0 Then
     psPat = Trim(Mid(lsNombre, 1, lnPos - 1))
     psPat = Trim(Replace(psPat, "-", ""))
     lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
     lnPos = InStr(1, lsNombre, "\")
     If lnPos > 0 Then
         psMat = Trim(Mid(lsNombre, 1, lnPos - 1))
         psMat = Trim(Replace(Replace(psMat, "-", ""), "\", " "))
         lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
         lnPos = InStr(1, lsNombre, ",")
         If lnPos > 0 Then
             psCas = Trim(Mid(lsNombre, 1, lnPos - 1))
             psCas = Trim(Replace(Replace(psCas, "-", ""), "\", " "))
             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
             lnPos = InStr(1, lsNombre, " ")
             If lnPos > 0 Then
                 psNom1 = Mid(lsNombre, 1, lnPos - 1)
                 psNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
             Else
                 psNom1 = Trim(lsNombre)
             End If
         End If
     Else
         lnPos = InStr(1, lsNombre, ",")
         If lnPos > 0 Then
             psMat = Mid(lsNombre, 1, lnPos - 1)
             psMat = Trim(Replace(Replace(psMat, "-", ""), "\", " "))
             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
             lnPos = InStr(1, lsNombre, " ")
             If lnPos > 0 Then
                 psNom1 = Mid(lsNombre, 1, lnPos - 1)
                 psNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
             Else
                 psNom1 = Trim(lsNombre)
             End If
         End If
     End If
 Else
     lnPos = InStr(1, lsNombre, "\")
     If lnPos > 0 Then
         psMat = Trim(Mid(lsNombre, 1, lnPos - 1))
         psMat = Trim(Replace(Replace(psMat, "-", ""), "\", " "))
         lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
         lnPos = InStr(1, lsNombre, ",")
         If lnPos > 0 Then
             psCas = Trim(Mid(lsNombre, 1, lnPos - 1))
             psCas = Trim(Replace(Replace(psCas, "-", ""), "\", " "))
             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
             lnPos = InStr(1, lsNombre, " ")
             If lnPos > 0 Then
                 psNom1 = Mid(lsNombre, 1, lnPos - 1)
                 psNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
             Else
                 psNom1 = Trim(lsNombre)
             End If
         End If
     Else
         lnPos = InStr(1, lsNombre, ",")
         If lnPos > 0 Then
             psMat = Mid(lsNombre, 1, lnPos - 1)
             psMat = Trim(Replace(Replace(psMat, "-", ""), "\", " "))
             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
             lnPos = InStr(1, lsNombre, " ")
             If lnPos > 0 Then
                 psNom1 = Mid(lsNombre, 1, lnPos - 1)
                 psNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
             Else
                 psNom1 = Trim(lsNombre)
             End If
         Else
             lsRazon = lsNombre
         End If
     End If
 End If
 
 If lsRazon <> "" Then
     psPat = Trim(lsRazon)
     psMat = ""
     psNom1 = ""
     psNom2 = ""
     psCas = ""
 Else
     psPat = Trim(IIf(psPat = "", "XXXX", psPat))
     psMat = Trim(IIf(psMat = "", "XXXX", psMat))
     If InStr(1, Trim(psCas), "DE LA ") = 0 Then
         lnPos = InStr(1, Trim(psCas), "DE ")
         If lnPos > 0 Then
             psCas = Trim(Mid(psCas, lnPos + 3, Len(psCas)))
         End If
     End If
 End If

 psPat = Replace(psPat, "Ñ", "#")
 psPat = Replace(psPat, "-", "")
 psPat = Replace(psPat, "|", "")
 psPat = Replace(psPat, ".", " ")
 psPat = Replace(psPat, "   ", " ")
 psPat = Replace(psPat, "  ", " ")
 
 psMat = Replace(psMat, "Ñ", "#")
 psMat = Replace(psMat, "-", "")
 psMat = Replace(psMat, "|", "")
 psMat = Replace(psMat, ".", " ")
 psMat = Replace(psMat, "   ", " ")
 psMat = Replace(psMat, "  ", " ")
 
 psCas = Replace(psCas, "Ñ", "#")
 psCas = Replace(psCas, "-", "")
 psCas = Replace(psCas, "|", "")
 psCas = Replace(psCas, ".", " ")
 psCas = Replace(psCas, "   ", " ")
 psCas = Replace(psCas, "  ", " ")
 
 psNom1 = Replace(psNom1, "Ñ", "#")
 psNom1 = Replace(psNom1, "-", "")
 psNom1 = Replace(psNom1, "|", "")
 psNom1 = Replace(psNom1, ".", " ")
 psNom1 = Replace(psNom1, "   ", " ")
 psNom1 = Replace(psNom1, "  ", " ")
 
 psNom2 = Replace(psNom2, "Ñ", "#")
 psNom2 = Replace(psNom2, "-", "")
 psNom2 = Replace(psNom2, "|", "")
 psNom2 = Replace(psNom2, ".", " ")
 psNom2 = Replace(psNom2, "   ", " ")
 psNom2 = Replace(psNom2, "  ", " ")

End Sub



Public Function GenerarDatosIBM(ByVal pbRespuesta As Boolean, _
                                ByVal pnMontoMinimoRCM As Double, _
                                ByVal pdFecDataFM As Date, _
                                ByVal pnTipCambio As Double, _
                                ByVal psServConsol As String, _
                                ByVal psMensaje As String) As Boolean

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim lsCadConexon As String
Dim CmdGenera As ADODB.Command
Dim Param1 As ADODB.Parameter

Dim contTotal  As Long
Dim J As Long
Dim lsApellPat As String
Dim lsApellMat As String
Dim lsApellCony As String
Dim lsNombre As String
Dim pos As Long
Dim Pos1 As Long
Dim Pos2 As Long
Dim Cont As Long
Dim lsPrefijos(6) As String
Dim i As Long

Dim lsCodPers As String
Dim lsCadAux As String

Dim lsTipoDoc As String
Dim lsNumDoc As String
Dim lsCalificacion As String
Dim lsTipoPers As String
Dim Total As Long
Dim lsNombrePers As String
Dim lvTipPers As String
Dim lvTidoCi As String
Dim lvNuDoCi  As String
Dim lvTidoTr As String
Dim lvNudoTr As String

On Error GoTo ErrorIbm
GenerarDatosIBM = False

If pbRespuesta Then

    oCon.ConexionActiva.BeginTrans

    Set CmdGenera = New ADODB.Command
    Set Param1 = New ADODB.Parameter

    CmdGenera.CommandText = psServConsol & "RCDGeneraIBM" '"RCDGeneraIBM"
    CmdGenera.CommandType = adCmdStoredProc
    CmdGenera.Name = "RCDGeneraIBM"
    Set Param1 = CmdGenera.CreateParameter("MontoMin", adInteger, adParamInput)
    CmdGenera.Parameters.Append Param1
    Set Param1 = CmdGenera.CreateParameter("Fecha", adDate, adParamInput)
    CmdGenera.Parameters.Append Param1
    Set Param1 = CmdGenera.CreateParameter("TipoCambio", adCurrency, adParamInput)
    CmdGenera.Parameters.Append Param1
    Set CmdGenera.ActiveConnection = oCon.ConexionActiva
    CmdGenera.CommandTimeout = 720
    'CmdGenera.Parameters.Refresh
    
   'ARCV 17-12-2006
   'oCon.ConexionActiva.RCDGeneraIBM pnMontoMinimoRCD, Format(pdFecDataFM, "mm/dd/yyyy"), pnTipCambio
   oCon.ConexionActiva.RCDGeneraIBM pnMontoMinimoRCM, Format(pdFecDataFM, "mm/dd/yyyy"), pnTipCambio
   '-------------
   oCon.ConexionActiva.CommitTrans
   
   Set CmdGenera = Nothing
   Set Param1 = Nothing
   
End If

Sql = "Select cCodPers, cApellPat, cApellMat, cNombre, cTipPers, cTipDoc, cNumDoc, cCalifica " & _
      "From " & psServConsol & "IBM1 WHERE cApellPat IS Null "

Set rs1 = oCon.CargaRecordSet(Sql)

If Not (rs1.BOF And rs1.EOF) Then
    Do While Not rs1.EOF
        Cont = Cont + 1
        '  P.cTidotr cTidoTr, P.cNudotr cNudoTr, P.cTidoci cTidoCi, P.cNudoci cNudoCi,
        Sql = "SELECT P.cPersCod,P.cPersNombre cNomPers, P.nPersPersoneria cTipPers,  " _
            & " R.cPersNom cNomPersRCD, R.cTipPers cTipPersRCD, R.cTidotr cTidoTrRCD, " _
            & " R.cNudotr cNudoTrRCD, R.cTidoci cTidoCiRCD, R.cNudoci cNudoCiRCD " _
            & "FROM Persona P LEFT JOIN " & psServConsol & "RCDMaestroPersona R " _
            & "ON P.cPersCod = R.cPersCod " _
            & "WHERE P.cPersCod='" & Trim(rs1!cCodpers) & "'  "
        Set rs = oCon.CargaRecordSet(Sql)
        
        If Not (rs.BOF And rs.EOF) Then
                lsApellPat = "": lsApellMat = "": lsNombre = "": lsApellCony = ""
                lsNombrePers = IIf(IsNull(rs!cNomPersRCD), "", Trim(rs!cNomPersRCD))
                If Len(Trim(lsNombrePers)) = 0 Then
                    lsNombrePers = Trim(rs!cNomPers)
                Else
                    If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
                        lsNombrePers = Trim(rs!cNomPers)
                    End If
                End If
                lsNombrePers = Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare))
                lsNombrePers = Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare))
                '-****Separa Nombre de Clientes
                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
                    If Trim(rs!cTipPers) = "1" Then
                        Pos1 = InStr(1, lsNombrePers, "/", vbTextCompare)
                        If Pos1 > 0 Then
                            lsApellPat = Mid(lsNombrePers, 1, Pos1 - 1)
                            lsCadAux = Mid(lsNombrePers, Pos1 + 1, Len(lsNombrePers))
                            
                            Pos1 = InStr(1, lsCadAux, "\", vbTextCompare)
                            If Pos1 > 0 Then
                                lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
                                lsCadAux = Mid(lsCadAux, Pos1 + 1, Len(lsCadAux))
                                If Len(Trim(lsApellMat)) = 0 Then
                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
                                    If Pos1 > 0 Then
                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
                                        lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
                                    Else
                                        lsApellCony = lsCadAux
                                    End If
                                Else
                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
                                    If Pos1 > 0 Then
                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
                                        lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
                                    Else
                                        lsApellCony = lsCadAux
                                    End If
                                    
                                End If
                            Else
                                Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
                                If Pos1 > 0 Then
                                    lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
                                    lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
                                End If
                            End If
                        Else
                            lsApellPat = lsNombrePers
                            lsApellMat = ""
                            lsNombre = ""
                        End If
                    Else
                        lsApellPat = lsNombrePers
                        lsApellMat = ""
                        lsNombre = ""
                    End If
                    If Len(Trim(lsApellCony)) > 0 Then
                        lsCadAux = lsNombre
                        lsNombre = lsApellPat & IIf(Len(Trim(lsApellMat)) = 0, " ", " " & Trim(lsApellMat) & " ") & "DE"
                        lsApellMat = lsCadAux
                        lsApellPat = lsApellCony
                    End If
                Else
                    lsApellPat = lsNombrePers
                    lsApellMat = ""
                    lsNombre = ""
                End If
                '-******
                lvTipPers = Trim(rs!cTipPers)
                
                Select Case lvTipPers
                    Case "1"
                        lsTipoPers = "1"
                        lsNumDoc = Trim(IIf(IsNull(lvNuDoCi), "", lvNuDoCi))
                        Select Case Trim(lvTidoCi)
                            Case "1"
                                lsTipoDoc = "1"
                            Case "3", "4" ' Carnet FFAA / FFPP
                                lsTipoDoc = "2"
                            Case "2", "5" ' Carnet Extranjeria / Pasaporte
                                lsTipoDoc = "3"
                        End Select
                    Case Else
                        If (InStr(1, lsApellPat, "Y/O", vbTextCompare) <> 0 Or InStr(1, lsApellPat, " O ", vbTextCompare) <> 0 Or InStr(1, lsApellPat, " Y ", vbTextCompare) <> 0) And (Len(Trim(lvNudoTr)) = 0) Then
                            lsTipoPers = "1"
                            lsNumDoc = Trim(IIf(IsNull(lvNuDoCi), "", lvNuDoCi))
                            Select Case Trim(lvTidoCi)
                                Case "1"
                                    lsTipoDoc = "1"
                                Case "3", "4" ' Carnet FFAA / FFPP
                                    lsTipoDoc = "2"
                                Case "2", "5" ' Carnet Extranjeria / Pasaporte
                                    lsTipoDoc = "3"
                            End Select
                        Else
                            lsTipoPers = "2"
                            lsNumDoc = Trim(IIf(IsNull(lvNudoTr), "", lvNudoTr))
                        End If
                        Select Case Trim(lvTidoTr)
                            Case "2", "3"
                                lsTipoDoc = "4"

                        End Select
                End Select

                '******************** CALIFICACION DE LA PERSONA **************
                lsCalificacion = nObtieneCalificacionPersonaProcesada(rs!cPersCod, psServConsol, psMensaje)
                
                If Trim(lsCalificacion) = "" Then
                    psMensaje = " OJO, no se esta considerando calificacion "
                End If
                
                'oCon.Ejecutar " RCDActualizaIBM '" & Replace(lsApellPat, "'", "''") & "','" & Replace(lsApellMat, "'", "''") & "','" & Replace(lsNombre, "'", "''") & "','" & Trim(lsCalificacion) & "','" & lsTipoDoc & "','" & lsNumDoc & "','" & lsTipoPers & "','" & Trim(rs1!cCodPers) & "'"
                oCon.Ejecutar " " & psServConsol & "RCDActualizaIBM '" & Replace(lsApellPat, "'", "''") & "','" & Replace(lsApellMat, "'", "''") & "','" & Replace(lsNombre, "'", "''") & "','" & Trim(lsCalificacion) & "','" & lsTipoDoc & "','" & lsNumDoc & "','" & lsTipoPers & "','" & Trim(rs1!cCodpers) & "'"
        End If
        rs.Close
        Set rs = Nothing
        
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
GenerarDatosIBM = True

Exit Function
ErrorIbm:
    oCon.ConexionActiva.RollbackTrans
    GenerarDatosIBM = False
    Err.Raise Err.Number, "Generar Datos IBM", Err.Description
    
End Function


'***********************************************************
'               FUNCIONES PARA EL MANEJO DE REPORTES
'***********************************************************

Private Function EliminaPunto(lnNumero As Currency) As Currency
Dim pos As Long
Dim CadAux As String
Dim CadAux1 As String
Dim lsNumero As String
lsNumero = Trim(Str(lnNumero))
If Val(lsNumero) > 0 Then
    pos = InStr(1, lsNumero, ".", vbTextCompare)
    If pos > 0 Then
        CadAux = Mid(lsNumero, 1, pos - 1)
        CadAux1 = Mid(lsNumero, pos + 1, Len(Trim(lsNumero)))
        If Len(Trim(CadAux1)) = 1 Then
            CadAux1 = CadAux1 & "0"
        End If
        EliminaPunto = CCur(CadAux & CadAux1)
    Else
        EliminaPunto = lnNumero & "00"
    End If
Else
    EliminaPunto = lnNumero
End If
End Function

Public Function nRepo179101_ArchivoRCD(ByVal psServerConsol As String, _
                                        ByVal pdFecSis As Date, _
                                        ByRef psMensaje As String) As String
Dim SQL1 As String
Dim Sql2 As String
Dim CadPrint() As String
Dim lsCadenaAux As String
Dim lsCadenaPrint As String
Dim TotalCadenas As Long
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String

Dim lsApellPat As String
Dim lsApellMat As String
Dim lsPrimerNombre As String
Dim lsSegundoNombre As String
Dim lsApellCony As String
Dim lsNombrePers As String
Dim Pos1 As Long
Dim Pos2 As Long
Dim Pos3 As Long
Dim lsCadAux As String
Dim lscadena As String
Dim lsTipPer As String
Dim lsGenero As String
Dim lsEstado As String
Dim lsCondCta As String
Dim lsCondDisp As String

'''
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSBS As String
Dim lsCalificacion As String
Dim lsCodPers As String

NúmeroArchivo = FreeFile

Dim lsPersonaAct As String
Dim lnSaldo As Currency

Dim lsNumSecICC As String
Dim lsIndRCC As String
' Crea el nombre de archivo.

Dim lsCodOfiInfICC As String
Dim lsUbicaGeoICC As String
Dim lsCodDptoICC As String

Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer

Dim lsImpre As String

'LAYG 2005/06/11 - Agregue el indicador de riesgo cambiario a Cero
Dim lsRiesCambiarioPersona As String
Dim lsCondEspCta As String
lsRiesCambiarioPersona = "0"

'LAYG 2006-07-15 - Indicador atraso, Calif Interna
Dim lsIndicadorAtr As String
Dim lsCalifInterna As String

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = oCon.CargaRecordSet(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

SQL1 = "select top 1 * from " & psServerConsol & "rcdparametro  order by  cMEs desc"
Set rs1 = oCon.CargaRecordSet(SQL1)
gnMontoMinimoICC = rs1!nMontoMin

lsNombreArchivo = App.Path & "\RCD" & Format(pdFecSis, "yyyymm") & ".106"
Open lsNombreArchivo For Output As #NúmeroArchivo

'**** Registro CABECERA
Print #NúmeroArchivo, "0106" & "01" & oCad.FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(pdFecSis, "yyyymmdd") & "012" _
        & Space(15) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False) & Chr(13) & Chr(10);
        

SQL1 = "SELECT * FROM " & psServerConsol & "sysobjects WHERE name = '" & "RCDvc" & Format(pdFecSis, "yyyymm") & "01'"
Set rs1 = oCon.CargaRecordSet(SQL1)

If rs1.EOF And rs1.BOF Then
    psMensaje = "Tabla [RCDvc" & Format(pdFecSis, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas"
    Exit Function
End If
      
SQL1 = " Select COUNT(*) AS TOTAL From " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 R1 " _
     & " Inner Join " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod "

Set rs1 = oCon.CargaRecordSet(SQL1)

If Not rs1.EOF Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

SQL1 = "Select R1.cTipoFor, R1.cTipoInf, R1.cNumSec, R1.cCodSBS, R1.cPersCod, R1.cCodUnico, R1.cActEcon, " _
    & "   R1.cCodRegPub, R1.cTidoTr, R1.cNudoTr, R1.cTiDoci, R1.cNuDoci, " _
    & "   R1.cTipPers, R1.cResid, R1.cCalifica, R1.cMagEmp, R1.cAccionista, R1.cRelInst,  " _
    & "   R1.cPaisNac, R1.cSiglas, R1.cPersNom,R1.cPersNomCom,R1.cPersGenero,R1.cPersEstado, R2.cTipoFor cTipoFor2, R2.cTipoInf cTipoInf2 , " _
    & "   R2.cCodAge, R2.cUbicGeo, R2.cCtaCnt, R2.cTipoCred, R2.nSaldo, R2.nCondDias,   " _
    & "   R2.cPersCod cCodPers2, cApePat, cApeMat, cApeCasada, cNombre1, cNombre2, " _
    & "   R2.cCondEspCta, R2.cCondDisponib, R1.cIndRCambiario, R1.cIndAtrasoDeudor, R1.cClasifInterna " _
    & " From " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 R1 " _
    & "      Inner Join " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod " _
    & " ORDER BY cApePat, cApeMat, cApeCasada, cNombre1, cNombre2 "
    

Set rs1 = oCon.CargaRecordSet(SQL1)
Cont = 0: J = 0

If Not rs1.EOF Then
    Do While Not rs1.EOF
        J = J + 1
            
        If lsPersonaAct <> rs1!cPersCod Then

            Cont = Cont + 1
            lsNumSecICC = oCad.FillNum(Trim(Str(Cont)), 6, "0")
            
            lsTipPer = rs1!cTipPers
            lsApellPat = "": lsApellMat = "": lsPrimerNombre = "": lsApellCony = "": lsSegundoNombre = ""
            lsNombrePers = rs1!cPersNomCom
            If Len(Trim(lsNombrePers)) = 0 Then
                    lsNombrePers = Trim(lsNombrePers)
            Else
                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
                   lsNombrePers = Trim(lsNombrePers)
                End If
            End If
            lsNombrePers = Trim(Replace(lsNombrePers, "-/", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "//", "/", , , vbTextCompare))
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not IsNull(rs1!cPersGenero) Then
                lsGenero = Trim(rs1!cPersGenero)
            Else
                lsGenero = ""
            End If
            If Not IsNull(rs1!cPersEstado) Then
                Select Case Trim(rs1!cPersEstado)
                    Case "2", "5", "6"
                        lsEstado = "C"
                    Case "3"
                        lsEstado = "V"
                    Case "4"
                         lsEstado = "D"
                    Case "1"
                         lsEstado = "S"
                    Case Else
                         lsEstado = Trim(rs1!cPersEstado)
                End Select
            Else
                lsEstado = ""
            End If

                        
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!cTidoci), "", IIf(Len(Trim(rs1!cTidoci)) = 0, "", Trim(rs1!cTidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cNudoci), "", IIf(Len(Trim(rs1!cNudoci)) = 0, "", Trim(rs1!cNudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidotr), "", IIf(Len(Trim(rs1!cTidotr)) = 0, "", Trim(rs1!cTidotr)))
                    lscNudotr = IIf(IsNull(rs1!cNudotr), "", IIf(Len(Trim(rs1!cNudotr)) = 0, "", Trim(rs1!cNudotr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!cCodSBS) Then
                lsCodSBS = Trim(rs1!cCodSBS)
            Else
                lsCodSBS = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = oCad.FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            
            'lsCodPers = EmiteCodigoPersona(rs1!cPersCod, psServerConsol)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cPersCod)
            'End If
            
            lsApellPat = Trim(rs1!cApePat)
            lsApellMat = Trim(rs1!cApeMat)
            lsApellCony = Trim(rs1!cApeCasada)
            lsPrimerNombre = Trim(rs1!cNombre1)
            lsSegundoNombre = Trim(rs1!cNombre2)
            lsIndRCC = Trim(rs1!cIndRCambiario)
            
            
            lsRiesCambiarioPersona = Trim(rs1!cIndRCambiario)
            
            lsIndicadorAtr = Trim(rs1!cIndAtrasoDeudor)
            lsCalifInterna = Trim(IIf(IsNull(rs1!cClasifInterna), "X", rs1!cClasifInterna))
            If Trim(lsCalifInterna) = "X" Then
                lsCalifInterna = "XXXXX"
            End If
            
            
            Print #NúmeroArchivo, rs1!cTipoFor & rs1!cTipoInf & oCad.FillNum(Trim(lsNumSecICC), 8, " ") & _
                    oImpre.ImpreFormat(IIf(Len(Trim(lsCodSBS)) = 0, "", lsCodSBS), 10, 0) & _
                    oImpre.ImpreFormat(Trim(rs1!cCodUnico), 20, 0) & oCad.FillNum(lsActividad, 4, " ") & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cCodRegPub), "", rs1!cCodRegPub)), 15, " ") & _
                    oCad.FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    oCad.FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & _
                    oCad.FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    oImpre.ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    oCad.FillNum(Trim(rs1!cCalifica), 1, " ") & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & _
                    oCad.FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    oCad.FillNum(Trim(rs1!cPaisNac), 4, " ") & oCad.FillNum(Trim(lsGenero), 1, " ") & _
                    oCad.FillNum(Trim(lsEstado), 1, " ") & _
                    oImpre.ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) & _
                    oImpre.ImpreFormat(Trim(lsApellPat), 120, 0) & _
                    oImpre.ImpreFormat(Trim(lsApellMat), 40, 0) & _
                    oImpre.ImpreFormat(Trim(lsApellCony), 40, 0) & _
                    oImpre.ImpreFormat(Trim(lsPrimerNombre), 40, 0) & _
                    oImpre.ImpreFormat(Trim(lsSegundoNombre), 40, 0) & oImpre.ImpreFormat(Trim(lsIndRCC), 1, 0) & _
                    oImpre.ImpreFormat(Trim(IIf(IsNull(rs1!cIndAtrasoDeudor), " ", rs1!cIndAtrasoDeudor)), 1, 0) & _
                    oImpre.ImpreFormat(Trim(IIf(IsNull(rs1!cClasifInterna), " ", rs1!cClasifInterna)), 5, 0); Chr(13) & Chr(10);
                    
        End If
        
        
        If Val(rs1!cCondEspCta) <= 3 Then
            lsCondEspCta = oCad.FillNum(Trim(rs1!cCondEspCta), 2, " ")
        Else
            Select Case lsRiesCambiarioPersona  ' 2006/03/15 susy
                Case "0"
                    lsCondEspCta = "13"
                Case "1"
                    lsCondEspCta = "11"
                Case "2"
                    lsCondEspCta = "12"
            End Select
        End If
        
        Print #NúmeroArchivo, rs1!cTipoFor2 & rs1!cTipoInf2 & oCad.FillNum(Trim(lsNumSecICC), 8, " ") & _
                oCad.FillNum(Trim(rs1!cCodAge), 4, "0") & oCad.FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                FormatoCtaContable(rs1!cCtaCnt) & oCad.FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                oCad.FillNum(Trim(rs1!nCondDias), 4, " ") & lsCondEspCta & _
                oCad.FillNum(Trim(rs1!cCondDisponib), 2, " "); Chr(13) & Chr(10);
        
        '********************* ACTUALIZANDO NUMERO DE SECUENCIA PARA REPORTE ICC *******************
        'SQL1 = "Update " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        'Sql2 = "Update " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        'oCon.Ejecutar (SQL1)
        'oCon.Ejecutar (Sql2)
        '*******************************************************************************************
        lsPersonaAct = rs1!cPersCod
        rs1.MoveNext
    Loop
    
End If
rs1.Close
Set rs1 = Nothing
'***************************************
'** Totales de la Empresa
'***************************************

Cont = Cont + 1
lsNumSecICC = oCad.FillNum(Trim(Str(Cont)), 6, "0")

Print #NúmeroArchivo, lsImpre & "2" & "1" & oCad.FillNum(Trim(lsNumSecICC), 6, " ") & _
        Space(10) & Chr(13) & Chr(10);

SQL1 = "Select * from " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "

Set rs1 = oCon.CargaRecordSet(SQL1)
If Not rs1.EOF Then
    Do While Not rs1.EOF
        
        Print #NúmeroArchivo, "2" & "2" & oCad.FillNum(Trim(lsNumSecICC), 8, " ") & _
                oCad.FillNum(Trim(lsCodOfiInfICC), 4, " ") & oCad.FillNum(Trim(lsUbicaGeoICC), 6, " ") & _
                FormatoCtaContable(rs1!cCtaCnt) & oCad.FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                oCad.FillNum(Trim(rs1!nCondDias), 4, " "); Chr(13) & Chr(10);
        
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing

Close #NúmeroArchivo   ' Cierra el archivo.
nRepo179101_ArchivoRCD = lsImpre

End Function


Public Function nRepo179103_IBM(ByVal psServerConsol As String, ByVal dFecSis As Date, ByVal psNomCmact As String, ByVal psCodCMAC As String) As String

Dim lsNombrePers As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim contTotal  As Long
Dim Sql As String
Dim J As Long
Dim lsApellPat As String
Dim lsApellMat As String
Dim lsApellCony As String
Dim lsNombre As String
Dim pos As Long
Dim Pos1 As Long
Dim Pos2 As Long
Dim Cont As Long
Dim lsTipoDoc As String
Dim lsNumDoc As String
Dim lsPrefijos(6) As String
Dim i As Long
Dim m As Long
Dim lsCodPers As String
Dim lsCadAux As String

Dim TOTALTIPODOC1 As Long: Dim TOTALTIPODOC2 As Long: Dim TOTALTIPODOC3 As Long
Dim TOTALTIPODOC4 As Long: Dim TOTALTIPODOC5 As Long: Dim TOTALTIPODOC6 As Long

Dim TOTALPERSNAT As Long: Dim TOTALPERSJUR As Long

Dim TOTALVENCMEN30MN As Long: Dim TOTALVENCMAY30MN As Long
Dim TOTALVENCMEN30ME As Long: Dim TOTALVENCMAY30ME As Long

Dim TOTALCAL0 As Long: Dim TOTALCAL1 As Long: Dim TOTALCAL2 As Long:
Dim TOTALCAL3 As Long: Dim TOTALCAL4 As Long

Dim TOTALNORMMN As Currency: Dim TOTALREFMN As Currency: Dim TOTALMEN30MN As Currency
Dim TOTALMAY30MN  As Currency: Dim TOTALCOBJUDMN  As Currency: Dim TOTALLINCREDMN As Currency:
Dim TOTALCASTMN As Currency
Dim lsCalificacion As String
Dim TOTALNORMME As Currency: Dim TOTALREFME As Currency: Dim TOTALMEN30ME As Currency
Dim TOTALMAY30ME  As Currency: Dim TOTALCOBJUDME  As Currency:  Dim TOTALLINCREDME As Currency
Dim TOTALCASTME As Currency
'------->
Dim gsNombreEntInfIBM As String
Dim gsNomCortoEntInfIBM As String
Dim gsCodEmpInfICC As String
Dim lsNumSecICC As String

Dim lsImpre As String

Sql = "select nConsSisValor from ConstSistema where nConsSisCod=80"
Set rs = oCon.CargaRecordSet(Sql)
gsCodEmpInfICC = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
gsNombreEntInfIBM = psNomCmact
gsNomCortoEntInfIBM = psCodCMAC
lsPrefijos(1) = "DE LA"
lsPrefijos(2) = "DE LOS"
lsPrefijos(3) = "DE "
lsPrefijos(4) = "LA "
lsPrefijos(5) = "DEL "
lsPrefijos(6) = "DI "


lsImpre = oCad.FillNum(Trim(gsCodEmpInfICC), 3, " ") & oCad.FillNum("1", 1, " ") & oCad.FillNum(Trim(Format(dFecSis, "YYYYMMDD")), 8, " ") & _
        Space(11) & oImpre.ImpreFormat(gsNombreEntInfIBM, 80, 0) & oImpre.ImpreFormat(Trim(gsNomCortoEntInfIBM), 10, 0) & _
        Space(546) & "X"
'------>

Sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) NOT IN('Ñ') "
Set rs = oCon.CargaRecordSet(Sql)


contTotal = rs!Total
rs.Close
Set rs = Nothing

Sql = "Select cCodPers,cApellPat,cApellMat,cNombre,cTipPers,cTipDoc,cNumDoc,cCalifica," _
& " nNormMN,nRefMN,nVen30MN,nVen31MN,nCobJudMN,nLinCredMN,nCastMN,nNormME,nRefME, " _
& " nVen30ME , nVen31ME, nCobJudME, nLinCredME, nCastME " _
& " From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) <> 'Ñ' order by cApellPat,cApellmat "

'Select * From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) not in('Ñ') order by cApellPat,cApellmat"

Set rs = oCon.CargaRecordSet(Sql)
Cont = 0
If Not rs.EOF Then
    Do While Not rs.EOF
        Cont = Cont + 1
        lsNumSecICC = oCad.FillNum(Trim(Str(Cont)), 6, "0")

        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        
        lsImpre = lsImpre & oCad.FillNum(Trim(gsCodEmpInfICC), 3, " ") & oCad.FillNum(Trim("2"), 1, " ") & oCad.FillNum(Trim(lsNumSecICC), 6, " ") & _
                    oCad.FillNum(Trim(rs!CTIPDOC), 1, " ") & oImpre.ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    oImpre.ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & oImpre.ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    oImpre.ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & oCad.FillNum(Trim(rs!cTipPers), 1, " ") & _
                    oImpre.ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

Sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) IN('Ñ') "
Set rs = oCon.CargaRecordSet(Sql)

contTotal = rs!Total
rs.Close
Set rs = Nothing

Sql = "Select * From " & psServerConsol & "IBM1 where SUBSTRING(CAPELLPAT,1,1) in('Ñ') order by cApellPat,cApellmat"
Set rs = oCon.CargaRecordSet(Sql)

m = 0
If Not rs.EOF Then
    Do While Not rs.EOF
        m = m + 1
        Cont = Cont + 1
        lsNumSecICC = oCad.FillNum(Trim(Str(Cont)), 6, "0")
        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        lsImpre = lsImpre & oCad.FillNum(Trim(gsCodEmpInfICC), 3, " ") & oCad.FillNum(Trim("2"), 1, " ") & oCad.FillNum(Trim(lsNumSecICC), 6, " ") & _
                    oCad.FillNum(Trim(rs!CTIPDOC), 1, " ") & oImpre.ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    oImpre.ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & oImpre.ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    oImpre.ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & oCad.FillNum(Trim(rs!cTipPers), 1, " ") & _
                    oImpre.ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
                    oImpre.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

Sql = " SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='1'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALPERSNAT = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='2'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALPERSJUR = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close


Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='1'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='2'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='3'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='4'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='5'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC5 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='6'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALTIPODOC6 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
TOTALVENCMEN30MN = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30MN>0"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALVENCMEN30MN = TOTALVENCMEN30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMEN30ME = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30ME>0"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALVENCMEN30ME = TOTALVENCMEN30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30MN = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31MN>0"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALVENCMAY30MN = TOTALVENCMAY30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30ME = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31Me>0"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALVENCMAY30ME = TOTALVENCMAY30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='0'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCAL0 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='1'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCAL1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='2'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCAL2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='3'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCAL3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='4'"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCAL4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT SUM(nNormMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALNORMMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nRefMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALREFMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen30MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALMEN30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen31MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALMAY30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCobJudMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCOBJUDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nLinCredMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALLINCREDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCastMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCASTMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

'monto soles
Sql = "SELECT SUM(nNormME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALNORMME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nRefME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALREFME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen30ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALMEN30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen31ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALMAY30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCobJudME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCOBJUDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nLinCredME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALLINCREDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCastME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = oCon.CargaRecordSet(Sql)
TOTALCASTME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

lsImpre = lsImpre & oCad.FillNum(Trim(gsCodEmpInfICC), 3, " ") & oCad.FillNum(Trim("3"), 1, " ") & _
            oImpre.ImpreFormat(TOTALTIPODOC1, 8, 0) & oImpre.ImpreFormat(TOTALTIPODOC2, 8, 0) & oImpre.ImpreFormat(TOTALTIPODOC3, 8, 0) & oImpre.ImpreFormat(TOTALTIPODOC4, 8, 0) & oImpre.ImpreFormat(TOTALTIPODOC5, 8, 0) & oImpre.ImpreFormat(TOTALTIPODOC6, 8, 0) & _
            oImpre.ImpreFormat(TOTALPERSNAT, 8, 0) & oImpre.ImpreFormat(TOTALPERSJUR, 8, 0) & _
            oImpre.ImpreFormat(TOTALVENCMEN30MN, 8, 0) & oImpre.ImpreFormat(TOTALVENCMAY30MN, 8, 0) & oImpre.ImpreFormat(TOTALVENCMEN30ME, 8, 0) & oImpre.ImpreFormat(TOTALVENCMAY30ME, 8, 0) & _
            oImpre.ImpreFormat(TOTALCAL0, 8, 0) & oImpre.ImpreFormat(TOTALCAL1, 8, 0) & oImpre.ImpreFormat(TOTALCAL2, 8, 0) & oImpre.ImpreFormat(TOTALCAL3, 8, 0) & oImpre.ImpreFormat(TOTALCAL4, 8, 0) & _
            oImpre.ImpreFormat(EliminaPunto(TOTALNORMMN), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALREFMN), 15, 0, False) & _
            oImpre.ImpreFormat(EliminaPunto(TOTALMEN30MN), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALMAY30MN), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALCOBJUDMN), 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALLINCREDMN), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALCASTMN), 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & _
            oImpre.ImpreFormat(EliminaPunto(TOTALNORMME), 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALREFME), 15, 0, False) & _
            oImpre.ImpreFormat(EliminaPunto(TOTALMEN30ME), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALMAY30ME), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALCOBJUDME), 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALLINCREDME), 15, 0, False) & oImpre.ImpreFormat(EliminaPunto(TOTALCASTME), 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & oImpre.ImpreFormat(0, 15, 0, False) & _
            oImpre.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)

End Function

Public Function GetServerConsol() As String
    Dim rs As New ADODB.Recordset
    Set rs = oCon.CargaRecordSet("SELECT nConsSisValor from ConstSistema where nConsSisCod=43")
    GetServerConsol = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
    Set rs = Nothing
End Function

'************* Para Verificacion de Datos*********************
Public Function nRepo179201_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "   Select cPersCod,cPersNom,cCalifica as Total " _
            & " from " & psServerConsol & "RCDvc" & psFecha & "01 where (cCalifica ='' or cCalifica is null or cCalifica not in('0','1','2','3','4'))"

Set rs = GetQuery(Sql)
Set nRepo179201_ = rs
Set rs = Nothing
End Function

Public Function nRepo179202_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom, cTipPers, cCodSBS from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (PATINDEX('%Y/O%', cPersNom )<>0  OR  PATINDEX('% O %', cPersNom )<>0 OR PATINDEX('% Y %', cPersNom)<>0 )  and cTipPers<>'3' " _
        & " ORDER BY cTipPers"
Set rs = GetQuery(Sql)
Set nRepo179202_ = rs
Set rs = Nothing
End Function

Public Function nRepo179203_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cCodSBS, cPersNom from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (cCodSbs IS NULL OR cCodSbs ='' or LEN(cCodSbs)<>10) "
Set rs = GetQuery(Sql)
Set nRepo179203_ = rs
Set rs = Nothing
End Function

Public Function nRepo179204_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoTr FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS ='2' AND CMAGEMP  NOT IN ('1','2','3','4') "
Set rs = GetQuery(Sql)
Set nRepo179204_ = rs
Set rs = Nothing
End Function

Public Function nRepo179205_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoci FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS NOT IN ('2') AND CMAGEMP IN ('1','2','3','4')"
Set rs = GetQuery(Sql)
Set nRepo179205_ = rs
Set rs = Nothing
End Function

Public Function nRepo179206_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp , cNudoci, cNudotr  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (CMAGEMP IS NULL  OR CMAGEMP ='' OR CMAGEMP NOT IN ('0','1','2','3','4'))"
Set rs = GetQuery(Sql)
Set nRepo179206_ = rs
Set rs = Nothing
End Function

Public Function nRepo179207_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidoci, cNudoci, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE cTidoci not in ('1','2','3','4','5') and cTipPers in ('1','3')"
Set rs = GetQuery(Sql)
Set nRepo179207_ = rs
Set rs = Nothing
End Function

Public Function nRepo179208_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cNudoci, cCodSbS FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE (cNudoCi like '%000000%' or cNudoci = '' or cNudoci is null)  and cTipPers in ('1','3')"
Set rs = GetQuery(Sql)
Set nRepo179208_ = rs
Set rs = Nothing
End Function

Public Function nRepo179209_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cNudoTr like '%000000%' or cNudoTr = '' or cNudoTr is null)  and cTipPers in ('2')"
Set rs = GetQuery(Sql)
Set nRepo179209_ = rs
Set rs = Nothing
End Function

Public Function nRepo179210_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidotr, cNudotr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE cTidotr not in ('2','3') and cTipPers in ('2')"
Set rs = GetQuery(Sql)
Set nRepo179210_ = rs
Set rs = Nothing
End Function

Public Function nRepo179211_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cTipPers not in ('1','2','3','4') or cTipPers is null or cTipPers='')"
Set rs = GetQuery(Sql)
Set nRepo179211_ = rs
Set rs = Nothing
End Function

Public Function nRepo179212_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs,cActEcon  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE len(cActeCon)<>4 or cActecon like '%000%' or cActecon ='' or cActecon is Null"
Set rs = GetQuery(Sql)
Set nRepo179212_ = rs
Set rs = Nothing
End Function

Public Function nRepo179213_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cSiglas is Not Null or cSiglas<>'') "
Set rs = GetQuery(Sql)
Set nRepo179213_ = rs
Set rs = Nothing
End Function

Public Function nRepo179214_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodRegPub, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cCodRegPub like '%00000000000%' or cCodRegPub is null or cCodRegPub ='' ) and cTipPers ='2'"
Set rs = GetQuery(Sql)
Set nRepo179214_ = rs
Set rs = Nothing
End Function

Public Function nRepo179215_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod, cPersNom   FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE  CTIPPERS IN ('1') ORDER BY cPersCod "
Set rs = GetQuery(Sql)
Set nRepo179215_ = rs
Set rs = Nothing
End Function

Public Function nRepo179216_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoTr)<>11 and cTidoTr='2' )  and cTipPers in ('2')  "
Set rs = GetQuery(Sql)
Set nRepo179216_ = rs
Set rs = Nothing
End Function

Public Function nRepo179217_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoCi)<>8 and cTidoci='1' )  and cTipPers in ('1','3')  "
Set rs = GetQuery(Sql)
Set nRepo179217_ = rs
Set rs = Nothing
End Function

Function GetQuery(ByVal Sql As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Set rs = oCon.CargaRecordSet(Sql)
Set GetQuery = rs
End Function
'****************************************************


Sub nCreaTablasRCM(ByVal lsFecha As String, ByVal psServConsol As String)
Dim Sql As String

If nExisteTabla("RCMvc" & lsFecha & "01", psServConsol) = False Then
    Sql = "CREATE TABLE " & psServConsol & "RCMvc" & lsFecha & "01 ( " _
        & "cTipoFor char (1) NOT NULL ,cTipoInf char (1) NOT NULL ," _
        & "cNumSec char (8) NULL , cCodSBS char (10) NULL, " _
        & "cPersCod char (13) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " _
        & "cCodUnico char(15) NOT NULL, " _
        & " cActEcon char (4) NULL , cCodRegPub char (15) NULL , " _
        & "cTidoTr char (1) NULL , cNudoTr char (11) NULL , " _
        & "cTiDoci char (1) NULL , cNuDoci char (12) NULL , " _
        & "cTipPers char (1) NULL , " _
        & "cResid char (1) NULL , " _
        & "cCalifica char (1) NULL , " _
        & "cMagEmp char (1) NULL , " _
        & "cAccionista char (1) NULL , " _
        & "cRelInst char (1) NULL , " _
        & "cPaisNac char (4) NULL, cSiglas char (20) NULL ," _
        & "cPersNom char (120) NULL ,cPersNomCom Varchar (120) NULL , " _
        & "cApePat Varchar (120) NULL ,  " _
        & "cApeMat Varchar (40) NULL , " _
        & "cApeCasada Varchar (40) NULL ," _
        & "cNombre1  VarChar (40) NULL, cNombre2  VarChar (40) NULL, " _
        & "cPersGenero char (1) NULL , cPersEstado char (1) NULL , " _
        & "cTitular char (13) NULL, cCalificaTit char(1) NULL ) " _
        '& "CONSTRAINT PK_RCDvc" & lsFecha & "01" & Format(Time, "hh") & "_4__10 PRIMARY KEY (cCodPers) )"
        oCon.Ejecutar (Sql)
     
    Sql = "CREATE  INDEX cPersCod ON dbo.RCMvc" & lsFecha & "01 (cPersCod)"
    'oCon.Ejecutar (sql)
Else
    Sql = "DELETE FROM " & psServConsol & "RCMvc" & lsFecha & "01"
    oCon.Ejecutar (Sql)

End If
End Sub

Sub nCreaTablasRCA(ByVal lsFecha As String, ByVal psServConsol As String)
Dim Sql As String

'*****************BEGIN Comentado by NAGL 20200710******************'
'If nExisteTabla("RCAvc" & lsFecha & "01", psServConsol) = False Then
'    Sql = "CREATE TABLE " & psServConsol & "RCAvc" & lsFecha & "01 ( " _
'        & "cTipoFor char (1) NOT NULL ,cTipoInf char (1) NOT NULL ," _
'        & "cNumSec char (8) NULL , cCodSBS char (10) NULL, " _
'        & "cPersCod char (13) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " _
'        & "cCodUnico char(15) NOT NULL,  " _
'        & "cActEcon char (4) NULL , cCodRegPub char (15) NULL , " _
'        & "cTidoTr char (1) NULL , cNudoTr char (11) NULL , " _
'        & "cTiDoci char (1) NULL , cNuDoci char (12) NULL , " _
'        & "cTipPers char (1) NULL , " _
'        & "cResid char (1) NULL , " _
'        & "cCalifica char (1) NULL , " _
'        & "cMagEmp char (1) NULL , " _
'        & "cAccionista char (1) NULL , " _
'        & "cRelInst char (1) NULL , " _
'        & "cPaisNac char (4) NULL, cSiglas char (20) NULL ," _
'        & "cPersNom Varchar (120) NULL ,cPersNomCom Varchar (120) NULL , " _
'        & "cApePat Varchar (120) NULL ,  " _
'        & "cApeMat Varchar (40) NULL , "
'    Sql = Sql & " " _
'        & "cApeCasada Varchar (40) NULL ," _
'        & "cNombre1  VarChar (40) NULL, cNombre2  VarChar (40) NULL, " _
'        & "cPersGenero char (1) NULL , cPersEstado char (1) NULL , " _
'        & "cTitular char (13) NULL, cCalificaTit char(1) NULL, " _
'        & "cCodAge char (4) NULL , cUbicGeo char (6) NULL , " _
'        & "cCuentaTit char(18) NULL, " _
'        & "nSKTit money NULL )" _
'        '& "CONSTRAINT PK_RCAvc" & lsFecha & "01" & Format(Time, "hh") & "_4__10 PRIMARY KEY (cCodPers) )"
'      oCon.Ejecutar (Sql)
'
'    'sql = "CREATE  INDEX cCodPers ON dbo.RCAvc" & lsFecha & "01 (cCodPers)"
'    'oCon.Ejecutar (sql)
'
'Else
'    Sql = "DELETE FROM " & psServConsol & "RCAvc" & lsFecha & "01"
'    oCon.Ejecutar (Sql)
'End If
'*****************END Comentado by NAGL 20200710******************'
'******Agregado by NAGL 20200703************'
Sql = "EXEC stp_ins_nCreaTablasRCA '" & lsFecha & "','" & psServConsol & "'"
oCon.Ejecutar (Sql)
'*************END NAGL 20200703*************'
End Sub

'----------------------------------------
'***********************************************************
'************ Genera los Datos para el RCA *****************
'***********************************************************
Public Sub GeneraDatosRCA(ByVal psFecha As String, ByVal psServConsol As String, _
                        ByVal pdFecDataFM As Date, ByVal pnTipoCambio As Double, _
                        ByVal psCodOfiInf As String, ByVal psUbicaGeoRCD As String, _
                        ByVal pnMontoMinimoRCD As Double, ByVal psCtaNoPref As String, _
                        ByRef psMensaje As Variant, ByRef psMensaje1 As Variant)

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim rsPers As ADODB.Recordset
Dim rsRCD As ADODB.Recordset

Dim lbEncuentroEnRCD  As Boolean
Dim loRCDProceso As COMNCredito.NCOMRCD

Dim lsRelacionInst  As String
Dim rsCred As New ADODB.Recordset

Dim i As Long, J As Long
Dim Agencia As String
Dim lscadena As String
Dim lsNombrePersona As String
Dim lsTipoDocRCD As String * 1
Dim lsTipoInfRCD As String * 1
Dim lsCodPers As String
Dim lsIndRCC  As String

Dim lsCredito As String
Dim lnSaldoCap As Currency
Dim lnCapVenc As Currency
Dim lnDiasAtraso As Integer
Dim lbCobJud As Boolean
Dim lbDemanda As Boolean
Dim lbCastigado As Boolean
Dim lbRefinan As Boolean
Dim lnIntDeveng As Currency
Dim lnProvision As Currency
Dim lnMontoGar1 As Currency  ' Garant Hipotecaria
Dim lnMontoGar2 As Currency  ' Otras Garantias

Dim lsEstadoCredito As String
Dim lsCuentaCnt As String
    
Dim rsGarant As ADODB.Recordset
    
'Datos usados a Nivel de Formulario
Dim lcCodUnico As String * 15  ' Codigo asignado por la empresa
Dim lcCodSBS As String * 10
Dim lcNomPers As String * 120
Dim lcNomPersComp As String * 120
Dim lcGenero As String * 1
Dim lcEstadoCiv As String * 1
Dim lcActEcon As String * 4
Dim lcCodRegPub As String * 15
Dim lcTiDoTr As String * 1
Dim lcNuDoTr As String * 11
Dim lcTiDoCi As String * 1
Dim lcNuDoCi As String * 12
Dim lcTipPers As String * 1
Dim lcResid As String * 1
Dim lcMagEmp As String * 1
Dim lcAccionista As String * 1
Dim lcRelInst As String * 1
Dim lcPaisNac As String * 4
Dim lcSiglas As String * 20
Dim lcCalificacion As String * 1

Dim lsPat  As String
Dim lsMat  As String
Dim lsCas  As String
Dim lsNom1 As String
Dim lsNom2 As String

Dim rsSaldosCont As New ADODB.Recordset
Dim lsCondEspecial As String
Dim lsOpeEnGarant As String
Dim lnDiasGarantia As Integer ' Dias desde ult TAsac Garant

Dim lcTitular As String * 13
Dim lcTitularCalifica As String * 1
Dim lcCuentaTit As String * 18
Dim lcSKTit As Double
Dim lsCuentaCnt_SP As String
Dim lnSaldoGar As Currency

'Dim gmAgSBS(38) As String
Dim gmAgSBS() As String 'EJVG20160709
'Dim gmAgUbicacionSBS(38) As String
Dim gmAgUbicacionSBS() As String 'EJVG20160709
'Obtiene equivalencias de Agencias
'lsSQL = "Select cAgecod, cAgenciaSBS, cUbiGeoSBS from agencias"
lsSQL = "Select cAgecod, cAgenciaSBS, cUbiGeoSBS, cAgeCodMax = (select max(cAgeCod) from Agencias) from agencias" 'EJVG20160709
Set rs = oCon.CargaRecordSet(lsSQL)
Do While Not rs.EOF
    'EJVG20160709 ***
    If rs.Bookmark = 1 Then
        ReDim gmAgSBS(rs!cAgeCodMax)
        ReDim gmAgUbicacionSBS(rs!cAgeCodMax)
    End If
    'END EJVG *******

    gmAgSBS(rs!cAgeCod) = rs!cAgenciaSBS
    gmAgUbicacionSBS(rs!cAgeCod) = rs!cUbiGeoSBS
    rs.MoveNext
Loop
rs.Close

'** Carga los Creditos (Vigentes)
'lsSQL = " SELECT C.cCtaCod, C.nPrdEstado, C.nDiasAtraso, C.cNumFuente, C.nMontoApr, " _
'    & "         C.nSaldoCap, ISNULL(C.nCapVencido,0) nCapVenc, C.cRefinan, C.dFecVig, " _
'    & "         ISNULL(C.nIntDev,0) nIntDeveng , ISNULL(C.nIntSusp, 0) nIntSusp, C.nDemanda, "
'    'ARCV 16-06-2007 :Comentar para el proximo mes (poner solo PP.cPersCod)
'    lsSQL = lsSQL & " PP.cPersCod ,"
'    lsSQL = lsSQL & " cPersCodAnt= ISNULL(RCLI.cCodCli,''),"
'    '---------------------
'    lsSQL = lsSQL & "         RMP.cCodUnico, RMP.cNomPers AS cPersNombre , RMP.cTidoci, " _
'    & "         RMP.cNudoci, RMP.cTidotr, RMP.cNudotr,  RMP.cTipPers, " _
'    & "         RMP.cCodSBS,  RMP.cActEcon, RMP.cCodRegPub, RMP.cResid, RMP.cMagEmp, RMP.cAccionista, RMP.cRelInst, " _
'    & "         RMP.cPaisNac, RMP.cSiglas,  RMP.cGenero, RMP.cEstadoCiv, " _
'    & "         RMP.cApePat, RMP.cApeMat, RMP.cApeCasada, RMP.cPriNom, RMP.cSegNom, " _
'    & "         isnull(Ti.Titular,'') Titular,isnull( Ti.CalificaTit,'') CalificaTit,Isnull( Ti.CreditoTit,'') CreditoTit,Isnull( Ti.SKTit,0) SkTit, " _
'    & "         cCalif =     (Select isnull(max(ca1.cCalGen),'') FROM ColocCalifProv ca1 WHERE ca1.cPersCod = PP.cPersCod ),  " _
'    & "         Age.cUbiGeoSBS, Age.cAgenciaSBS "
'lsSQL = lsSQL & " " _
'    & " FROM " & psServConsol & "CreditoConsol C " _
'    & " INNER JOIN " & psServConsol & "ProductoPersonaConsol PP ON PP.cCtaCod = C.cCtaCod " _
'    & " LEFT JOIN " & psServConsol & "RCDMaestroPersona RMP ON RMP.cPersCod = PP.cPersCod " _
'    & " INNER JOIN Agencias Age ON Age.cAgeCod=SUBSTRING(C.cCtaCod,4,2) " _
'    & " INNER JOIN (Select ca.cCtaCod CreditoTit, ca.cPersCod as Titular, ca.cCalGen as CalificaTit, ca.nMontoApr SKTit " _
'    & "            From ColocCalifProv ca WHERE ca.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) ) Ti " _
'    & " ON Ti.CreditoTit = C.cCtaCod "
    
'& " inner join " & psServConsol & "Rcavc2007110101 Ant On Ant.cCuentaTit=C.cCtaCod "
'By Capi 1312007 solo por este mes, luego sacar
'By Capi 13122007 se modifico para que jale el monto aprobado en vez del saldo de capital ca.nSaldoCap por ca.nMontoApr
'ARCV 16-06-2007 : Comentar para el proximo mes
'lsSQL = lsSQL & " LEFT JOIN RelClientes RCLI ON Ti.Titular =  RCLI.cPersCod " 'ARCV 16-06-2007
''----------
'lsSQL = lsSQL & " WHERE C.nPrdEstado in( " & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor _
'    & "                       ," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor _
'    & "                       ," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov _
'    & "                        ) " _
'    & "AND PP.nPrdPersRelac = 25 AND C.nSaldoCap > " & Trim(pnMontoMinimoRCD) _
'    & " AND PP.cPersCod = (Select max(p1.cPersCod) From " & psServConsol & "ProductoPersonaConsol p1 " _
'    & "                where p1.cCtacod = c.cCtaCod and p1.nPrdPersRelac = 25  ) "
lsSQL = "exec stp_sel_RecuperaDatosParaGenerarRCA '" & Format(pdFecDataFM, "YYYY/MM/DD") & "'"
'Solo los vigentes ( se quito los judiciales)

i = 0

lsTipoDocRCD = "1"
lsTipoInfRCD = "1"

 Set rs = oCon.CargaRecordSet(lsSQL)
 
 'By Capi 13122007 para que elimine a los avales del RCD02,RCD02T Y RCD03
 lsSQL = "   DELETE FROM " & psServConsol & "RCDvc" & psFecha & "02T" _
            & "   WHERE cCtaCnt In ('8414100000','8424100000') "
            
 oCon.Ejecutar (lsSQL)
 
 lsSQL = "   DELETE FROM " & psServConsol & "RCDvc" & psFecha & "02" _
            & "   WHERE cCtaCnt In ('8414100000','8424100000') "
 
 oCon.Ejecutar (lsSQL)
 
 lsSQL = "   DELETE FROM " & psServConsol & "RCDvc" & psFecha & "03" _
            & "   WHERE cCtaCnt In ('8414100000','8424100000') "
 
 oCon.Ejecutar (lsSQL)
 'End By
 

 ReDim psMensaje(0)
 ReDim psMensaje1(0)
 
 If Not (rs.BOF And rs.EOF) Then
    Do While Not rs.EOF
        i = i + 1
        '***********************************************
        '**  DATOS DE LA PERSONA  -  RCDvcAAAAMM01   ***
        '***********************************************
        '***** limpio las variables
        lcCodUnico = "":         lcCodSBS = ""
        lcNomPers = "":          lcActEcon = ""
        lcNomPersComp = ""
        lcCodRegPub = "":        lcTiDoTr = ""
        lcNuDoTr = "":           lcTiDoCi = ""
        lcNuDoCi = "":           lcTipPers = ""
        lcResid = "":            lcMagEmp = ""
        lcAccionista = "":       lcRelInst = ""
        lcPaisNac = "":          lcSiglas = ""
        lcCalificacion = "0"
        lcGenero = ""
        lcEstadoCiv = ""
        
        lsCodPers = Trim(EmiteCodigoAux(Trim(rs!cPersCod), psServConsol))
        If Len(Trim(lsCodPers)) = 0 Then
            lsCodPers = Trim(rs!cPersCod)
        End If
        
        lcTitular = rs!Titular
        
        'ARCV 20-06-2007 : Cambiar al codigo antiguo (Solo para Mayo)
        '**Comentado por DAOR 20070913
        'If Trim(rs!cPersCodAnt) <> "" Then
        '    lcTitular = Trim(rs!cPersCodAnt)
        'End If
        '------------
        
        lcTitularCalifica = rs!CalificaTit
        lcCuentaTit = rs!CreditoTit
        
        lcSKTit = rs!SKTit
        lsCuentaCnt_SP = rs!cCtaContCod
        lnSaldoGar = rs!nMontoGar
        'Verifica si se encuentra en el RCA
        'By Capi 13122007 comentado porque puede existir mas de un garante a distintos deudores
        'lsSQL = "Select cPersCod From " & psServConsol & "RCAvc" & Format(pdFecDataFM, "yyyymm") & "01 " _
        '      & " Where cPersCod ='" & Trim(lsCodPers) & "'"
                        
        'Set rsRCD = oCon.CargaRecordSet(lsSQL)
        'lbEncuentroEnRCD = IIf((rsRCD.BOF And rsRCD.EOF) = True, False, True)
            lbEncuentroEnRCD = False
        'Set rsRCD = Nothing
        
        If lbEncuentroEnRCD = False Then ' Lo inserta
            If Not (IsNull(rs!cCodUnico)) Then ' Existe en Maestro Persona
                
                ' Asigno los valores a la Variables de RCDMaestroPersona
                lcCodUnico = rs!cCodUnico
                lcCodSBS = IIf(IsNull(rs!cCodSBS), "", rs!cCodSBS)
                lcNomPers = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
                lcNomPers = fReemplazaCaracterEspecial(lcNomPers)
                'lcNomPersComp = fReemplazaCaracterEspecial(rs!cPersNombre)
                
                lcActEcon = IIf(IsNull(rs!cActEcon), "", rs!cActEcon)
                lcCodRegPub = IIf(IsNull(rs!cCodRegPub), "", rs!cCodRegPub)
                lcTiDoTr = IIf(IsNull(rs!cTidotr), "", rs!cTidotr)
                lcNuDoTr = IIf(IsNull(rs!cNudotr), "", rs!cNudotr)
                lcTiDoCi = IIf(IsNull(rs!cTidoci), "", rs!cTidoci)
                lcNuDoCi = IIf(IsNull(rs!cNudoci), "", rs!cNudoci)
                lcTipPers = rs!cTipPers
                If lcTipPers = "3" Then
                   lcTipPers = "2"
                End If
                lcResid = rs!cResid
                lcMagEmp = IIf(IsNull(rs!cMagEmp), "", rs!cMagEmp)
                lcAccionista = rs!cAccionista
                lcRelInst = rs!cRelInst
                lcPaisNac = rs!cPaisNac
                lcSiglas = IIf(IsNull(rs!cSiglas), "", rs!cSiglas)
                lcGenero = IIf(IsNull(rs!cGenero), "", rs!cGenero)
                lcEstadoCiv = IIf(IsNull(rs!cEstadoCiv), "", rs!cEstadoCiv)
                
                lsPat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApePat), "", rs!cApePat)))
                lsMat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeMat), "", rs!cApeMat)))
                lsCas = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeCasada), "", rs!cApeCasada)))
                lsNom1 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cPriNom), "", rs!cPriNom)))
                lsNom2 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cSegNom), "", rs!cSegNom)))
                '***
            Else ' Busco en Personas
                '***** LUCV20170411, Comentó y agregó
'                lsSQL = "SELECT P.cPersCod, P.cPersNombre, P.dPersNacCreac, ISNULL(P.nPersPersoneria,'') As cTipPers, " _
'                        & "IsNULL(P.cPersCIIU,'') AS cActEcon, ISNULL(P.nPersRelaInst,'') As cRelInst,  " _
'                        & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
'                        & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
'                        & " ISNULL(PJ.cPersJurMagnitud,'') As cMagEmp,ISNULL(PN.cPersNatMagnitud,'') As cMagPersNat,  ISNULL(PJ.cPersJurSigla,'') As cSiglas,  " _
'                        & " IsNULL(PN.cPersNatSexo,'0') AS Genero,  IsNULL(PN.nPersNatEstCiv,'0') AS cEstadoCiv  "
'                        'ARCV 16-06-2007
'                        'jaca 20110428 se agrego el campo ISNULL(PN.cPersNatMagnitud,'') As cMagPersNat
'                        lsSQL = lsSQL & " ,cCodSBS=ISNULL(P.cPersCodSBS,'') "
'                        '--------
'                        lsSQL = lsSQL & " From Persona P " _
'                            & " LEFT JOIN PersonaJur PJ on P.cPersCod = PJ.cPersCod " _
'                            & " LEFT JOIN PersonaNat PN on P.cPersCod = PN.cPersCod " _
'                            & " WHERE P.cPersCod ='" & lsCodPers & "' "
                            
                        lsSQL = "exec stp_sel_RecuperaDatosCliente '" & lsCodPers & "'"
                '***** Fin LUCV20170411
                
                Set rsPers = oCon.CargaRecordSet(lsSQL)

                If rsPers.BOF And rsPers.EOF Then
                    MsgBox "NO ENCUENTRO LA PERSONA", vbInformation, "Aviso"
                Else
                    'lcCodUnico = lsCodPers
                    
                    'ARCV 16-06-2007
                    'lcCodSBS = ""
                    lcCodSBS = rsPers!cCodSBS
                    '--------

                    lcNomPersComp = rsPers!cPersNombre
                    lcNomPers = fReemplazaCaracterEspecial(rsPers!cPersNombre)
                    CambiaNombreRCD rsPers!cPersNombre, lsPat, lsMat, lsCas, lsNom1, lsNom2
                    lsPat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsPat), "", lsPat)))
                    lsMat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsMat), "", lsMat)))
                    lsCas = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsCas), "", lsCas)))
                    lsNom1 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsNom1), "", lsNom1)))
                    lsNom2 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsNom2), "", lsNom2)))
                    lcActEcon = Right(rsPers!cActEcon, 4)
                    lcTipPers = rsPers!cTipPers
                    lcMagEmp = rsPers!cMagEmp
                    lcRelInst = rsPers!cRelInst
                    lcSiglas = rsPers!cSiglas
                    
                    If lcTipPers <> "1" Then
                        lcNuDoTr = IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                        lcTiDoTr = IIf(IsNull(rsPers!NroRUC), "", 2)
                        lcTiDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                        lcNuDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                        lcCodUnico = lcNuDoTr
                    Else
                        lcNuDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                        lcTiDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", 2)
                        lcTiDoCi = Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                        lcNuDoCi = Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                        lcCodUnico = lcNuDoCi
                        
                        lcCodRegPub = ""
                        'lcMagEmp = "0"
                        lcMagEmp = rsPers!cMagPersNat 'jaca 20110428
                        lcSiglas = ""
                    End If
                    If lcTipPers = "3" Then
                        lcTipPers = "2"
                    End If

                    lcResid = "1"

                    If rsPers!cRelInst = "A" Then
                        lcAccionista = "1"
                    Else
                        lcAccionista = "0"
                    End If

                    If Not IsNull(rsPers!cRelInst) Then
                       Select Case Trim(rsPers!cRelInst)
                        Case "N", "O"  'ninguno/ no indicado
                            lcRelInst = "0"
                        Case "D"        'director
                            lcRelInst = "1"
                        Case "F"        'funcionario
                            lcRelInst = "2"
                        Case "T"        'trabajador
                            lcRelInst = "3"
                        Case "A"
                            lcRelInst = "0"
                        End Select
                    Else
                        lcRelInst = "0"
                    End If
                    'ALPA 20120206
                    'lcPaisNac = "4028"
                    lcPaisNac = "PE"
                    lcGenero = rsPers!Genero
                    lcEstadoCiv = "S"

                    Select Case rsPers!cEstadoCiv
                        Case Null: lcEstadoCiv = "0"
                        Case "1": lcEstadoCiv = "S" '1
                        Case "2": lcEstadoCiv = "C" '2
                        Case "3": lcEstadoCiv = "V" '4
                        Case "4": lcEstadoCiv = "D" '3
                        Case "5": lcEstadoCiv = "C" '2
                        Case "6": lcEstadoCiv = "S" '1
                    End Select
                    If lcEstadoCiv = "D" And Len(lsCas) > 0 Then
                        lcEstadoCiv = "C"
                    End If
                End If
                rsPers.Close
            End If
            
            '******************** CALIFICACION DE LA PERSONA **************
            lcCalificacion = " " 'Vacia
            'Saldo se convierte a Soles
            'lcSKTit = IIf(Mid(lcCuentaTit, 9, 1) = "1", lcSKTit, Round(lcSKTit * pnTipoCambio, 2))

            'ARCV 16-06-2007 : Comentar para el proximo mes
            '**Comentado por DAOR 20070912
            'If Trim(rs!cPersCodAnt) <> "" Then
            '    lsCodPers = rs!cPersCodAnt
            'End If
            '------

            lsSQL = "INSERT INTO " & psServConsol & "RCAvc" & psFecha & "01" _
                & "(cTipoFor,cTipoInf,cNumSec, cCodSBS,cPersCod, cCodUnico, cActEcon, cCodRegPub, " _
                & " cTidoTr, cNudoTr, cTiDoci, cNuDoci, cTipPers, cResid, cCalifica, cMagEmp, " _
                & " cAccionista, cRelInst,cPaisNac, cSiglas,cPersNom, cPersGenero,cPersEstado, " _
                & " cApePat, cApeMat ,cApeCasada, cNombre1, cNombre2, cTitular, cCalificaTit, " _
                & " cCodAge, cUbicGeo, cCuentaTit, nSKTit ) " _
                & " VALUES('" & lsTipoDocRCD & "','" & lsTipoInfRCD & "',Null, " _
                & "'" & lcCodSBS & "','" & lsCodPers & "','" & lcCodUnico & "', " _
                & "'" & lcActEcon & "'" & "," & "'" & lcCodRegPub & "'" & "," _
                & IIf(lcTiDoTr = "Null", "Null", "'" & Trim(lcTiDoTr) & "'") & "," _
                & IIf(lcNuDoTr = "Null", "Null", "'" & Trim(lcNuDoTr) & "'") & ",'" _
                & Trim(IIf(IsNull(lcTiDoCi), "1", lcTiDoCi)) & "'," _
                & Trim(IIf(IsNull(lcNuDoCi), "NULL", "'" & Trim(lcNuDoCi) & "'")) & ",'" _
                & Trim(lcTipPers) & "','" _
                & lcResid & "','" & lcCalificacion & "','" & lcMagEmp & "','" _
                & lcAccionista & "','" & lcRelInst & "','" & lcPaisNac & "'," _
                & "'" & lcSiglas & "'" & ",'" _
                & Trim(lcNomPers) & "','" & lcGenero & "','" & lcEstadoCiv & "','" _
                & Trim(lsPat) & "','" & Trim(lsMat) & "','" & Trim(lsCas) & "','" & Trim(lsNom1) & "','" & Trim(lsNom2) & "','" _
                & lcTitular & "','" & lcTitularCalifica & "','" _
                & rs!cAgenciaSBS & "','" & rs!cUbiGeoSBS & "','" & lcCuentaTit & "'," & lcSKTit & " ) "
                
            oCon.Ejecutar (lsSQL)
            
            'By Capi 13122007 para que inserte en el RCD la cuenta 8414100000 o 8424100000 segun corresponda
            If Len(Trim(lcTitular)) > 0 Then
                lsCuentaCnt = "84" & Mid(lsCuentaCnt_SP, 3, 1) & "4100000"
                lsCodPers = lcTitular
                lnMontoGar2 = lnSaldoGar
                lsCondEspecial = "04"
                lsOpeEnGarant = "02"
            'Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnMontoGar2, Mid(lcCuentaTit, 6, 1), pnTipoCambio, lsCondEspecial, lsOpeEnGarant, psServConsol, gmAgSBS(Mid(lcCuentaTit, 4, 2)), rs!cUbiGeoSBS, lcCuentaTit)
            Call GrabaEnRCDSaldos(psFecha, lsCodPers, lsCuentaCnt, 0, lnMontoGar2, rs!cTpoCreditoEquivalente, pnTipoCambio, lsCondEspecial, lsOpeEnGarant, psServConsol, gmAgSBS(rs!cAgeCodB), rs!cUbiGeoSBS, lcCuentaTit)
            End If
            'End By
        End If
        
        rs.MoveNext
    Loop
End If
rs.Close

'By Capi 13122007 para que actualice en detalle de saldos
' Actualiza el Total en RCDvcXXXXX02 a partir de RCDvcXXXXX02T
lsSQL = "INSERT INTO " & psServConsol & "RCDvc" & psFecha & "02 (cTipoFor,cTipoInf,cNumSec, cCodAge," _
      & " cUbicGeo, cCtaCnt, cTipoCred,nSaldo, nCondDias, cCondEspCta, cCondDisponib, cPersCod, cCtaCod) " _
      & " SELECT cTipoFor, cTipoInf, null, cCodAge, cUbicgeo, cCtaCnt, ctipoCred, " _
      & " sum(nsaldo), nCondDias , cCondEspCta, cCondDisponib, cPersCod, cCtaCod " _
      & " From " & psServConsol & "RCDvc" & psFecha & "02T " _
      & " Where cCtaCnt in ('8414100000','8424100000') " _
      & " Group By cTipoFor, cTipoInf, cCodAge, cUbicgeo, cCtaCnt, cTipoCred, nCondDias, cCondEspCta, cCondDisponib, cPersCod, cCtaCod  "
oCon.Ejecutar lsSQL

' Actualiza el Total en RCDvcXXXXX03
lsSQL = "INSERT INTO " & psServConsol & "RCDvc" & psFecha & "03 (cTipoFor,cTipoInf,cNumSec,cCtaCnt, cTipoCred,nSaldo, " _
      & " nCondDias, cCondEspCta, cCondDisponib) " _
      & " Select '2','2',Null, cCtaCnt, cTipoCred, Sum(nSaldo), nCondDias, cCondEspCta, max(cConddisponib) " _
      & " From " & psServConsol & "RCDvc" & psFecha & "02 " _
      & " Where cCtaCnt in ('8414100000','8424100000') " _
      & " Group By cCtaCnt, cTipoCred, nCondDias, cCondEspCta " _
      & " Order By cCtaCnt, cTipocred, nconddias, cCondEspCta "
oCon.Ejecutar lsSQL

'End By











Set rs = Nothing
 
End Sub

'***********************************************************
'************ Genera los Datos para el RCM *****************
'***********************************************************
Public Sub GeneraDatosRCM(ByVal psFecha As String, ByVal psServConsol As String, _
                        ByVal pdFecDataFM As Date, ByVal pnTipoCambio As Double, _
                        ByVal psCodOfiInf As String, ByVal psUbicaGeoRCD As String, _
                        ByVal pnMontoMinimoRCD As Double, ByVal psCtaNoPref As String, _
                        ByRef psMensaje As Variant, ByRef psMensaje1 As Variant)

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim rsPers As ADODB.Recordset
Dim rsRCD As ADODB.Recordset

Dim lbEncuentroEnRCD  As Boolean
Dim loRCDProceso As COMNCredito.NCOMRCD

Dim lsRelacionInst  As String
Dim rsCred As New ADODB.Recordset

Dim i As Long, J As Long
Dim Agencia As String
Dim lscadena As String
Dim lsNombrePersona As String
Dim lsTipoDocRCD As String * 1
Dim lsTipoInfRCD As String * 1
Dim lsCodPers As String
Dim lsIndRCC  As String

Dim lsCredito As String
Dim lnSaldoCap As Currency
Dim lnCapVenc As Currency
Dim lnDiasAtraso As Integer
Dim lbCobJud As Boolean
Dim lbDemanda As Boolean
Dim lbCastigado As Boolean
Dim lbRefinan As Boolean
Dim lnIntDeveng As Currency
Dim lnProvision As Currency
Dim lnMontoGar1 As Currency  ' Garant Hipotecaria
Dim lnMontoGar2 As Currency  ' Otras Garantias

Dim lsEstadoCredito As String
Dim lsCuentaCnt As String
    
Dim rsGarant As ADODB.Recordset
    
'Datos usados a Nivel de Formulario
Dim lcCodUnico As String * 15  ' Codigo asignado por la empresa
Dim lcCodSBS As String * 10
Dim lcNomPers As String * 120
Dim lcNomPersComp As String * 120
Dim lcGenero As String * 1
Dim lcEstadoCiv As String * 1
Dim lcActEcon As String * 4
Dim lcCodRegPub As String * 15
Dim lcTiDoTr As String * 1
Dim lcNuDoTr As String * 11
Dim lcTiDoCi As String * 1
Dim lcNuDoCi As String * 12
Dim lcTipPers As String * 1
Dim lcResid As String * 1
Dim lcMagEmp As String * 1
Dim lcAccionista As String * 1
Dim lcRelInst As String * 1
Dim lcPaisNac As String * 4
Dim lcSiglas As String * 20
Dim lcCalificacion As String * 1

Dim lsPat  As String
Dim lsMat  As String
Dim lsCas  As String
Dim lsNom1 As String
Dim lsNom2 As String

Dim rsSaldosCont As New ADODB.Recordset
Dim lsCondEspecial As String
Dim lsOpeEnGarant As String
Dim lnDiasGarantia As Integer ' Dias desde ult TAsac Garant

Dim lcTitular As String * 13
Dim lcTitularCalifica As String * 1
Dim lcCuentaTit As String * 18
Dim lcSKTit As Double


Dim gmAgSBS(38) As String
Dim gmAgUbicacionSBS(38) As String
'Obtiene equivalencias de Agencias
lsSQL = "Select cAgecod, cAgenciaSBS, cUbiGeoSBS from agencias"
Set rs = oCon.CargaRecordSet(lsSQL)
Do While Not rs.EOF
    gmAgSBS(rs!cAgeCod) = rs!cAgenciaSBS
    gmAgUbicacionSBS(rs!cAgeCod) = rs!cUbiGeoSBS
    rs.MoveNext
Loop
rs.Close

'** Carga los Creditos (Vigentes)
lsSQL = " SELECT C.cCtaCod, C.nPrdEstado, C.nDiasAtraso, C.cNumFuente, C.nMontoApr, " _
    & "         C.nSaldoCap, ISNULL(C.nCapVencido,0) nCapVenc, C.cRefinan, C.dFecVig, " _
    & "         ISNULL(C.nIntDev,0) nIntDeveng , ISNULL(C.nIntSusp, 0) nIntSusp, C.nDemanda, " _
    'ARCV 16-06-2007 :Comentar para el proximo mes (poner solo PP.cPersCod)
    lsSQL = lsSQL & " PP.cPersCod ,"
    lsSQL = lsSQL & " cPersCodAnt= ISNULL(RCLI.cPersCod,''),"
    '---------------------
lsSQL = lsSQL & " RMP.cCodUnico, RMP.cNomPers AS cPersNombre , RMP.cTidoci, " _
    & "         RMP.cNudoci, RMP.cTidotr, RMP.cNudotr,  RMP.cTipPers, " _
    & "         RMP.cCodSBS,  RMP.cActEcon, RMP.cCodRegPub, RMP.cResid, RMP.cMagEmp, RMP.cAccionista, RMP.cRelInst, " _
    & "         RMP.cPaisNac, RMP.cSiglas,  RMP.cGenero, RMP.cEstadoCiv, " _
    & "         RMP.cApePat, RMP.cApeMat, RMP.cApeCasada, RMP.cPriNom, RMP.cSegNom, " _
    & "         Ti.Titular, Ti.CalificaTit, Ti.CreditoTit, Ti.SKTit, " _
    & "         cCalif = (Select isnull(max(ca1.cCalGen),'') FROM ColocCalifProv ca1 WHERE ca1.cPersCod = PP.cPersCod ),  " _
    & "         Age.cUbiGeoSBS, Age.cAgenciaSBS "
lsSQL = lsSQL & " " _
    & " FROM " & psServConsol & "CreditoConsol C " _
    & " INNER JOIN " & psServConsol & "ProductoPersonaConsol PP ON PP.cCtaCod = C.cCtaCod " _
    & " LEFT JOIN " & psServConsol & "RCDMaestroPersona RMP ON RMP.cPersCod = PP.cPersCod " _
    & " INNER JOIN Agencias Age ON Age.cAgeCod=SUBSTRING(C.cCtaCod,4,2) " _
    & " LEFT JOIN (Select ca.cCtaCod CreditoTit, ca.cPersCod as Titular, ca.cCalGen as CalificaTit, ca.nSaldoCap SKTit " _
    & "            From ColocCalifProv ca WHERE ca.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) ) Ti " _
    & " ON Ti.CreditoTit = C.cCtaCod "
    'ARCV 16-06-2007 : Comentar para el proximo mes
    lsSQL = lsSQL & " LEFT JOIN RelClientes RCLI ON PP.cPersCod =  RCLI.cPersCod " 'ARCV 16-06-2007
    '----------
lsSQL = lsSQL & " WHERE C.nPrdEstado in( " & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor _
    & "                       ," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor _
    & "                       ," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov _
    & "                        ) " _
    & "AND PP.nPrdPersRelac = 21 AND C.nSaldoCap > " & Trim(pnMontoMinimoRCD) _
    & " AND PP.cPersCod = (Select max(p1.cPersCod) From " & psServConsol & "ProductoPersonaConsol p1 " _
    & "                where p1.cCtacod = c.cCtaCod and p1.nPrdPersRelac = 21  ) "

'Solo los vigentes ( se quito los judiciales)

i = 0

lsTipoDocRCD = "1"
lsTipoInfRCD = "1"

 Set rs = oCon.CargaRecordSet(lsSQL)

 ReDim psMensaje(0)
 ReDim psMensaje1(0)
 
 If Not (rs.BOF And rs.EOF) Then
    Do While Not rs.EOF
        i = i + 1
        '***********************************************
        '**  DATOS DE LA PERSONA  -  RCDvcAAAAMM01   ***
        '***********************************************
        '***** limpio las variables
        lcCodUnico = "":         lcCodSBS = ""
        lcNomPers = "":          lcActEcon = ""
        lcNomPersComp = ""
        lcCodRegPub = "":        lcTiDoTr = ""
        lcNuDoTr = "":           lcTiDoCi = ""
        lcNuDoCi = "":           lcTipPers = ""
        lcResid = "":            lcMagEmp = ""
        lcAccionista = "":       lcRelInst = ""
        lcPaisNac = "":          lcSiglas = ""
        lcCalificacion = "0"
        lcGenero = ""
        lcEstadoCiv = ""
        
        lsCodPers = Trim(EmiteCodigoAux(Trim(rs!cPersCod), psServConsol))
        If Len(Trim(lsCodPers)) = 0 Then
            lsCodPers = Trim(rs!cPersCod)
        End If
        
        lcTitular = rs!Titular
        lcTitularCalifica = rs!CalificaTit
        lcCuentaTit = rs!CreditoTit
        lcSKTit = rs!SKTit
        
        'Verifica si se encuentra en el RCA
        lsSQL = "Select cPersCod From " & psServConsol & "RCMvc" & Format(pdFecDataFM, "yyyymm") & "01 " _
              & " Where cPersCod ='" & Trim(lsCodPers) & "'"
        
        Set rsRCD = oCon.CargaRecordSet(lsSQL)
            lbEncuentroEnRCD = IIf((rsRCD.BOF And rsRCD.EOF) = True, False, True)
        Set rsRCD = Nothing
        
        If lbEncuentroEnRCD = False Then ' Lo inserta
            If Not (IsNull(rs!cCodUnico)) Then ' Existe en Maestro Persona
                
                ' Asigno los valores a la Variables de RCDMaestroPersona
                lcCodUnico = rs!cCodUnico
                lcCodSBS = IIf(IsNull(rs!cCodSBS), "", rs!cCodSBS)
                lcNomPers = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
                lcNomPers = fReemplazaCaracterEspecial(lcNomPers)
                'lcNomPersComp = fReemplazaCaracterEspecial(rs!cPersNombre)
                
                lcActEcon = IIf(IsNull(rs!cActEcon), "", rs!cActEcon)
                lcCodRegPub = IIf(IsNull(rs!cCodRegPub), "", rs!cCodRegPub)
                lcTiDoTr = IIf(IsNull(rs!cTidotr), "", rs!cTidotr)
                lcNuDoTr = IIf(IsNull(rs!cNudotr), "", rs!cNudotr)
                lcTiDoCi = IIf(IsNull(rs!cTidoci), "", rs!cTidoci)
                lcNuDoCi = IIf(IsNull(rs!cNudoci), "", rs!cNudoci)
                lcTipPers = rs!cTipPers
                If lcTipPers = "3" Then
                   lcTipPers = "2"
                End If
                lcResid = rs!cResid
                lcMagEmp = IIf(IsNull(rs!cMagEmp), "", rs!cMagEmp)
                lcAccionista = rs!cAccionista
                lcRelInst = rs!cRelInst
                lcPaisNac = rs!cPaisNac
                lcSiglas = IIf(IsNull(rs!cSiglas), "", rs!cSiglas)
                lcGenero = IIf(IsNull(rs!cGenero), "", rs!cGenero)
                lcEstadoCiv = IIf(IsNull(rs!cEstadoCiv), "", rs!cEstadoCiv)
                
                lsPat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApePat), "", rs!cApePat)))
                lsMat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeMat), "", rs!cApeMat)))
                lsCas = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cApeCasada), "", rs!cApeCasada)))
                lsNom1 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cPriNom), "", rs!cPriNom)))
                lsNom2 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(rs!cSegNom), "", rs!cSegNom)))
                '***
            Else ' Busco en Personas
                lsSQL = "SELECT P.cPersCod, P.cPersNombre, P.dPersNacCreac, ISNULL(P.nPersPersoneria,'') As cTipPers, " _
                        & "IsNULL(P.cPersCIIU,'') AS cActEcon, ISNULL(P.nPersRelaInst,'') As cRelInst,  " _
                        & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
                        & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
                        & " ISNULL(PJ.cPersJurMagnitud,'') As cMagEmp,  ISNULL(PJ.cPersJurSigla,'') As cSiglas,  " _
                        & " IsNULL(PN.cPersNatSexo,'0') AS Genero,  IsNULL(PN.nPersNatEstCiv,'0') AS cEstadoCiv  "
                        'ARCV 16-06-2007
                        lsSQL = lsSQL & " ,cCodSBS=ISNULL(P.cPersCodSBS,'') "
                        '--------
                        lsSQL = lsSQL & " From Persona P " _
                            & " LEFT JOIN PersonaJur PJ on P.cPersCod = PJ.cPersCod " _
                            & " LEFT JOIN PersonaNat PN on P.cPersCod = PN.cPersCod " _
                            & " WHERE P.cPersCod ='" & lsCodPers & "' "

                Set rsPers = oCon.CargaRecordSet(lsSQL)

                If rsPers.BOF And rsPers.EOF Then
                    MsgBox "NO ENCUENTRO LA PERSONA", vbInformation, "Aviso"
                Else
                    'lcCodUnico = lsCodPers
                    'lcCodSBS = ""
                    lcCodSBS = rsPers!cCodSBS '"" 'ARCV 16-06-2007

                    lcNomPersComp = rsPers!cPersNombre
                    lcNomPers = fReemplazaCaracterEspecial(rsPers!cPersNombre)
                    CambiaNombreRCD rsPers!cPersNombre, lsPat, lsMat, lsCas, lsNom1, lsNom2
                    lsPat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsPat), "", lsPat)))
                    lsMat = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsMat), "", lsMat)))
                    lsCas = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsCas), "", lsCas)))
                    lsNom1 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsNom1), "", lsNom1)))
                    lsNom2 = fReemplazaCaracterEspecial(VerApostrofe(IIf(IsNull(lsNom2), "", lsNom2)))
                    lcActEcon = Right(rsPers!cActEcon, 4)
                    lcTipPers = rsPers!cTipPers
                    lcMagEmp = rsPers!cMagEmp
                    lcRelInst = rsPers!cRelInst
                    lcSiglas = rsPers!cSiglas
                    
                    If lcTipPers <> "1" Then
                        lcNuDoTr = IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                        lcTiDoTr = IIf(IsNull(rsPers!NroRUC), "", 2)
                        lcTiDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                        lcNuDoCi = "" ' Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                        lcCodUnico = lcNuDoTr
                    Else
                        lcNuDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
                        lcTiDoTr = "" ' IIf(IsNull(rsPers!NroRUC), "", 2)
                        lcTiDoCi = Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
                        lcNuDoCi = Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
                        lcCodUnico = lcNuDoCi
                        
                        lcCodRegPub = ""
                        lcMagEmp = "0"
                        lcSiglas = ""
                    End If
                    If lcTipPers = "3" Then
                        lcTipPers = "2"
                    End If

                    lcResid = "1"

                    If rsPers!cRelInst = "A" Then
                        lcAccionista = "1"
                    Else
                        lcAccionista = "0"
                    End If

                    If Not IsNull(rsPers!cRelInst) Then
                       Select Case Trim(rsPers!cRelInst)
                        Case "N", "O"  'ninguno/ no indicado
                            lcRelInst = "0"
                        Case "D"        'director
                            lcRelInst = "1"
                        Case "F"        'funcionario
                            lcRelInst = "2"
                        Case "T"        'trabajador
                            lcRelInst = "3"
                        Case "A"
                            lcRelInst = "0"
                        End Select
                    Else
                        lcRelInst = "0"
                    End If

                    lcPaisNac = "4028"
                    lcGenero = rsPers!Genero
                    lcEstadoCiv = "S"

                    Select Case rsPers!cEstadoCiv
                        Case Null: lcEstadoCiv = "0"
                        Case "1": lcEstadoCiv = "S" '1
                        Case "2": lcEstadoCiv = "C" '2
                        Case "3": lcEstadoCiv = "V" '4
                        Case "4": lcEstadoCiv = "D" '3
                        Case "5": lcEstadoCiv = "C" '2
                        Case "6": lcEstadoCiv = "S" '1
                    End Select
                    If lcEstadoCiv = "D" And Len(lsCas) > 0 Then
                        lcEstadoCiv = "C"
                    End If
                End If
                rsPers.Close
            End If
            
            '******************** CALIFICACION DE LA PERSONA **************
            lcCalificacion = IIf(IsNull(rs!cCalif), "", rs!cCalif)
            If lcCalificacion = " " Then ' si no tiene calificacion le asigna la del titular
                lcCalificacion = lcTitularCalifica
            End If
            
            'ARCV 16-06-2007 : Comentar para el proximo mes
            If Trim(rs!cPersCodAnt) <> "" Then
                lsCodPers = rs!cPersCodAnt
            End If
            '------
            
            
            lsSQL = "INSERT INTO " & psServConsol & "RCMvc" & psFecha & "01  " _
                & "(cTipoFor,cTipoInf,cNumSec, cCodSBS,cPersCod, cCodUnico, cActEcon, cCodRegPub, " _
                & " cTidoTr, cNudoTr, cTiDoci, cNuDoci, cTipPers, cResid, cCalifica, cMagEmp, " _
                & " cAccionista, cRelInst,cPaisNac, cSiglas,cPersNom, cPersGenero,cPersEstado, " _
                & " cApePat, cApeMat ,cApeCasada, cNombre1, cNombre2, cTitular, cCalificaTit) " _
                & " VALUES('" & lsTipoDocRCD & "','" & lsTipoInfRCD & "',Null, " _
                & "'" & lcCodSBS & "','" & lsCodPers & "','" & lcCodUnico & "', " _
                & "'" & lcActEcon & "'" & "," & "'" & lcCodRegPub & "'" & "," _
                & IIf(lcTiDoTr = "Null", "Null", "'" & Trim(lcTiDoTr) & "'") & "," _
                & IIf(lcNuDoTr = "Null", "Null", "'" & Trim(lcNuDoTr) & "'") & ",'" _
                & Trim(IIf(IsNull(lcTiDoCi), "1", lcTiDoCi)) & "'," _
                & Trim(IIf(IsNull(lcNuDoCi), "NULL", "'" & Trim(lcNuDoCi) & "'")) & ",'" _
                & Trim(lcTipPers) & "','" _
                & lcResid & "','" & lcCalificacion & "','" & lcMagEmp & "','" _
                & lcAccionista & "','" & lcRelInst & "','" & lcPaisNac & "'," _
                & "'" & lcSiglas & "'" & ",'" _
                & Trim(lcNomPers) & "','" & lcGenero & "','" & lcEstadoCiv & "','" _
                & Trim(lsPat) & "','" & Trim(lsMat) & "','" & Trim(lsCas) & "','" & Trim(lsNom1) & "','" & Trim(lsNom2) & "','" _
                & lcTitular & "','" & lcTitularCalifica & "' ) "
                
            oCon.Ejecutar (lsSQL)
        End If
        
        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing
 
End Sub

Public Function ObtenerVc(ByVal pdFecha As Date, psNumSec As String, ByVal psNomTabla As String, ByVal psTipo As String) As ADODB.Recordset
'JUEZ 20150310 Se agregó psNomTabla
 Dim oConec As COMConecta.DCOMConecta
 Dim rs As ADODB.Recordset
 Set oConec = New COMConecta.DCOMConecta
 
 Dim Sql As String
 
 Sql = " Select * "
 'Sql = Sql & " From dbconsolidada..RCDvc" & Format(pdFecha, "yyyymm") & IIf(psTipo = "01", "01", "02") & " R1 "
 Sql = Sql & " From dbconsolidada.." & psNomTabla & Format(pdFecha, "yyyymm") & IIf(psTipo = "01", "01", "02") & " R1 " 'JUEZ 20150310
 Sql = Sql & " where cNumSec = '" & psNumSec & "'"
 oConec.AbreConexion
 Set rs = oConec.CargaRecordSet(Sql)
 oConec.CierraConexion
 Set ObtenerVc = rs
End Function

Public Sub ActualizarVc01(pdFecha As Date, psNumSec As String, psNomTabla As String, _
                                lsMagEmp As String, _
                                ldFecNac As Date, _
                                lsTiDociComp As String, _
                                lsNuDociComp As String, _
                                lsPersNom As String, _
                                lsApePat As String, _
                                lsApeMat As String, _
                                lsApeCasada As String, _
                                lsNombre1 As String, _
                                lsNombre2 As String, lsCodSBS As String, _
                                lsRelInst As String, _
                                lsSexo As String)
                                'JUEZ 20140410 Se agregó lsSexo
                                'JUEZ 20150310 Se agregó psNomTabla
 Dim oConec As COMConecta.DCOMConecta
 Dim rs As ADODB.Recordset
 Set oConec = New COMConecta.DCOMConecta
 
 Dim Sql As String
    'Sql = " Update dbconsolidada..RCDvc" & Format(pdFecha, "yyyymm") & "01 "
    Sql = " Update dbconsolidada.." & psNomTabla & Format(pdFecha, "yyyymm") & "01 " 'JUEZ 20150310
    Sql = Sql & " Set "
    Sql = Sql & " cMagEmp = '" & lsMagEmp & "',"
    Sql = Sql & " dFecNac = '" & Format(ldFecNac, "YYYY/MM/DD") & "',"
    Sql = Sql & " cTiDociComp = '" & lsTiDociComp & "',"
    Sql = Sql & " cNuDociComp = '" & Trim(lsNuDociComp) & "',"
    Sql = Sql & " cPersNom = '" & lsPersNom & "',"
    Sql = Sql & " cPersNomCom = '" & lsPersNom & "',"
    Sql = Sql & " cApePat = '" & lsApePat & "',"
    Sql = Sql & " cApeMat = '" & lsApeMat & "',"
    Sql = Sql & " cApeCasada = '" & lsApeCasada & "',"
    Sql = Sql & " cNombre1 = '" & lsNombre1 & "',"
    Sql = Sql & " cNombre2 = '" & lsNombre2 & "',"
    Sql = Sql & " cCodSBS = '" & lsCodSBS & "',"
    Sql = Sql & " cRelInst = '" & lsRelInst & "'," 'ALPA 20120615
    Sql = Sql & " cPersGenero = '" & lsSexo & "'" 'JUEZ 20140410
    Sql = Sql & " where cNumSec = '" & psNumSec & "'"
 oConec.AbreConexion
 Set rs = oConec.CargaRecordSet(Sql)
 oConec.CierraConexion
End Sub

Public Function ActualizarVc02(pdFecha As Date, psNumSec As String, psNomTabla As String, _
                                lsAgeCod As String, lsUbigeo As String)
                                'JUEZ 20150310 Se agregó psNomTabla
 Dim oConec As COMConecta.DCOMConecta
 Dim rs As ADODB.Recordset
 Set oConec = New COMConecta.DCOMConecta
 
 Dim Sql As String
    'Sql = " Update dbconsolidada..RCDvc" & Format(pdFecha, "yyyymm") & "02 "
    Sql = " Update dbconsolidada.." & psNomTabla & Format(pdFecha, "yyyymm") & "02 " 'JUEZ 20150310
    Sql = Sql & " Set "
    Sql = Sql & " cCodAge = '" & lsAgeCod & "',"
    Sql = Sql & " cUbicGeo = '" & lsUbigeo & "'"
    Sql = Sql & " where cNumSec = '" & psNumSec & "'"
 oConec.AbreConexion
 Set rs = oConec.CargaRecordSet(Sql)
 oConec.CierraConexion
End Function

Public Function ObtenerUbigeoRCD() As ADODB.Recordset
 Dim oConec As COMConecta.DCOMConecta
 Dim rs As ADODB.Recordset
 Set oConec = New COMConecta.DCOMConecta
 
 Dim Sql As String
    Sql = "exec stp_sel_UbigeoRCD"
 oConec.AbreConexion
 Set rs = oConec.CargaRecordSet(Sql)
 oConec.CierraConexion
 Set ObtenerUbigeoRCD = rs
End Function

'***************BEGIN NAGL 202007 - ACTA N°046-2020*********************'
'Public Sub ObtenerDatosParaGenerarRCDvc_Tvc(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, ByVal pnMontoMinimoRCD As Double, Optional ByVal pbRCDTransf As Boolean = False)
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Public Sub ObtenerDatosParaGenerarRCDvc_Tvc(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, ByVal pnMontoMinimoRCD As Double, Optional ByVal pbRCDTransf As Integer)
 Dim oConec As COMConecta.DCOMConecta
 Dim Sql As String
 On Error GoTo ErrObtenerDatosParaGenerarRCDvc_Tvc
 Set oConec = New COMConecta.DCOMConecta
 oConec.AbreConexion
 'ObtieneDatosRCDvc_Tvc
 'Sql = "EXEC stp_sel_ObtieneDatosParaGenerarRCDvc_Tvc '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pnMontoMinimoRCD & ", " & IIf(pbRCDTransf, 1, 0) & " "
 'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 Sql = "EXEC stp_sel_ObtieneDatosParaGenerarRCDvc_Tvc '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pnMontoMinimoRCD & ", " & pbRCDTransf & " "
 oCon.Ejecutar (Sql)
 'Garantias RCD
 'Sql = "EXEC stp_sel_ObtieneGarantiasRCD '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & " "
  'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
   If pbRCDTransf <> 2 Then
    'Garantias
 Sql = "EXEC stp_sel_ObtieneGarantiasRCD '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & " "
 oCon.Ejecutar (Sql)
 End If
 
' Sql = "EXEC stp_sel_ObtieneGarantiasRCD '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & " "
' oCon.Ejecutar (Sql)
 If pbRCDTransf = 0 Then
    'Diferidos
    Sql = "EXEC stp_sel_ReporteInteresDiferidoAmpliados "
    oCon.Ejecutar (Sql)
 End If
 oConec.CierraConexion
 Set oConec = Nothing
 Exit Sub
ErrObtenerDatosParaGenerarRCDvc_Tvc:
    Err.Raise Err.Number, "Data RCDvc_Tvc ", Err.Description
End Sub

'Public Sub GeneraDatosRCD_New(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, ByRef psMensaje As Variant, Optional ByVal pbRCDTransf As Boolean = False)
Public Sub GeneraDatosRCD_New(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, ByRef psMensaje As Variant, Optional ByVal pbRCDTransf As Integer)
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 Dim oConec As COMConecta.DCOMConecta
 Dim Sql As String
 Dim lsNomTabla As String
 Set oConec = New COMConecta.DCOMConecta
 ReDim psMensaje(0)
 On Error GoTo ErrGeneraDatosRCD_New
 oConec.AbreConexion
   ' lsNomTabla = IIf(pbRCDTransf, "RCDTvc", "RCDvc")
   'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
   
    Select Case pbRCDTransf
    Case "1"
        lsNomTabla = "RCDTvc"
    Case "0"
        lsNomTabla = "RCDvc"
    Case "2"
        lsNomTabla = "RCDHvc"
    End Select
   
    ReDim Preserve psMensaje(5)
    'Genera Datos RCD_01
    psMensaje(0) = GeneraDatosRCD_Detalle(psFecha, psServConsol, lsNomTabla, pnTipoCambio, "RCD_01", pbRCDTransf)
    psMensaje(1) = GeneraDatosRCD_Detalle(psFecha, psServConsol, lsNomTabla, pnTipoCambio, "RCD_02", pbRCDTransf)
    psMensaje(2) = GeneraDatosRCD_Detalle(psFecha, psServConsol, lsNomTabla, pnTipoCambio, "RCD_02_Garant", pbRCDTransf)
    psMensaje(3) = GeneraDatosRCD_Detalle(psFecha, psServConsol, lsNomTabla, pnTipoCambio, "RCD_02_IntDif", pbRCDTransf)
    psMensaje(4) = GeneraDatosRCD_Detalle(psFecha, psServConsol, lsNomTabla, pnTipoCambio, "RCD_02_OtrosProc", pbRCDTransf)
    Call GeneraDatosRCD_Totales(psFecha, psServConsol, pnTipoCambio, pbRCDTransf)
    
 oConec.CierraConexion
 Set oConec = Nothing
 Exit Sub
ErrGeneraDatosRCD_New:
    Err.Raise Err.Number, "Proceso RCDvc_Tvc ", Err.Description
End Sub

'Public Function GeneraDatosRCD_Detalle(ByVal psFecha As String, ByVal psServConsol As String, ByVal lsNomTabla As String, ByVal pnTipoCambio As Double, ByVal psTipo As String, Optional ByVal pbRCDTransf As Boolean = False) As String
Public Function GeneraDatosRCD_Detalle(ByVal psFecha As String, ByVal psServConsol As String, ByVal lsNomTabla As String, ByVal pnTipoCambio As Double, ByVal psTipo As String, Optional ByVal pbRCDTransf As Integer) As String
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc

 Dim oConec As COMConecta.DCOMConecta
 Dim psSql As String
 Dim psMensaje As String
 Dim oRs As ADODB.Recordset
 Set oConec = New COMConecta.DCOMConecta
 Set oRs = New ADODB.Recordset
 oConec.AbreConexion
 
 If psTipo = "RCD_01" Then
   'psSql = "Exec stp_ins_GeneraDatosRCD_01 '" & psFecha & "','" & psServConsol & "', '" & lsNomTabla & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ""
    psSql = "Exec stp_ins_GeneraDatosRCD_01 '" & psFecha & "','" & psServConsol & "', '" & lsNomTabla & "', " & pnTipoCambio & ", " & pbRCDTransf & ""
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    
 ElseIf psTipo = "RCD_02" Then
    'psSql = "Exec stp_ins_GeneraDatosRCD_02T '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ""
    psSql = "Exec stp_ins_GeneraDatosRCD_02T '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & ""
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    
 ElseIf psTipo = "RCD_02_Garant" Then
    'psSql = "Exec stp_ins_GeneraDatosRCD_02T_Garantias '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ""
    psSql = "Exec stp_ins_GeneraDatosRCD_02T_Garantias '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & ""
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    
 ElseIf psTipo = "RCD_02_IntDif" Then
    'psSql = "Exec stp_ins_GeneraDatosRCD_02T_IntDif '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ""
    psSql = "Exec stp_ins_GeneraDatosRCD_02T_IntDif '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & ""
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    
 Else 'RCD_02_OtrosProc
    'psSql = "Exec stp_ins_GeneraDatosRCD_02T_OtrosProc '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ""
    psSql = "Exec stp_ins_GeneraDatosRCD_02T_OtrosProc '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & ""
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    
 End If
 
 Set oRs = oConec.CargaRecordSet(psSql)
 If Not oRs.BOF And Not oRs.EOF Then
    psMensaje = IIf(IsNull(oRs!cMensaje), "", oRs!cMensaje)
 Else
    psMensaje = ""
 End If
 GeneraDatosRCD_Detalle = psMensaje
 oConec.CierraConexion
 Set oConec = Nothing
End Function

'Public Sub GeneraDatosRCD_Totales(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, Optional ByVal pbRCDTransf As Boolean = False, Optional ByVal psCondicion As String = "")
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Public Sub GeneraDatosRCD_Totales(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, Optional ByVal pbRCDTransf As Integer, Optional ByVal psCondicion As String = "")
 Dim oConec As COMConecta.DCOMConecta
 Dim psSql As String
 Set oConec = New COMConecta.DCOMConecta
 oConec.AbreConexion
 'psSql = "EXEC stp_ins_GeneraDatosRCD_02_03 '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & IIf(pbRCDTransf, 1, 0) & ", '" & psCondicion & "'"
 'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 psSql = "EXEC stp_ins_GeneraDatosRCD_02_03 '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ", " & pbRCDTransf & ", '" & psCondicion & "'"
 oCon.Ejecutar (psSql)
 oConec.CierraConexion
 Set oConec = Nothing
End Sub

Public Sub ObtenerDatosParaGenerarRCA(ByVal psFecha As String)
 Dim oConec As COMConecta.DCOMConecta
 Dim Sql As String
 Set oConec = New COMConecta.DCOMConecta
 oConec.AbreConexion
 'ObtieneDatosRCA
 Sql = "EXEC stp_sel_RecuperaDatosParaGenerarRCA '" & psFecha & "'"
 oCon.Ejecutar (Sql)
 oConec.CierraConexion
 Set oConec = Nothing
End Sub

Public Sub GeneraDatosRCA_New(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double, ByRef psMensaje As Variant)
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Set oConec = New COMConecta.DCOMConecta
ReDim psMensaje(0)
oConec.AbreConexion
ReDim Preserve psMensaje(1)
psMensaje(0) = GeneraDatosRCA_Detalle(psFecha, psServConsol, pnTipoCambio)
Call GeneraDatosRCD_Totales(psFecha, psServConsol, pnTipoCambio, , "''8414100000'',''8424100000''")
oConec.CierraConexion
Set oConec = Nothing
End Sub

Public Function GeneraDatosRCA_Detalle(ByVal psFecha As String, ByVal psServConsol As String, ByVal pnTipoCambio As Double) As String
Dim oConec As COMConecta.DCOMConecta
Dim psSql As String
Dim psMensaje As String
Dim oRs As ADODB.Recordset
Set oConec = New COMConecta.DCOMConecta
Set oRs = New ADODB.Recordset
oConec.AbreConexion
'Procesa Data para RCA, y para el RCD_02,02T,y03 en la cuenta 840410
psSql = "Exec stp_ins_GeneraDatosRCA_01 '" & psFecha & "','" & psServConsol & "', " & pnTipoCambio & ""
Set oRs = oConec.CargaRecordSet(psSql)

If Not oRs.BOF And Not oRs.EOF Then
   psMensaje = IIf(IsNull(oRs!cMensaje), "", oRs!cMensaje)
Else
   psMensaje = ""
End If
GeneraDatosRCA_Detalle = psMensaje
oConec.CierraConexion
Set oConec = Nothing
End Function

'Public Sub GeneraSecuenciaRCD_RCA(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Boolean = False)
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Public Sub GeneraSecuenciaRCD_RCA(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Integer)

 Dim oConec As COMConecta.DCOMConecta
 Dim Sql As String
 Set oConec = New COMConecta.DCOMConecta
 oConec.AbreConexion
 'Genera Secuencia RCD_RCA
 'Sql = "EXEC stp_upd_GeneraSecuenciaRCD_RCA '" & psFecha & "','" & psServConsol & "'," & IIf(pbRCDTransf, 1, 0) & ""
 'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 Sql = "EXEC stp_upd_GeneraSecuenciaRCD_RCA '" & psFecha & "','" & psServConsol & "'," & pbRCDTransf & ""
 oCon.Ejecutar (Sql)
 'Elimina Secuencia
 'Sql = "EXEC stp_del_EliminaSecuenciaRCD '" & psFecha & "','" & psServConsol & "'," & IIf(pbRCDTransf, 1, 0) & ""
 'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 Sql = "EXEC stp_del_EliminaSecuenciaRCD '" & psFecha & "','" & psServConsol & "'," & pbRCDTransf & ""
 oCon.Ejecutar (Sql)
 'Genera Secuencia
 'Sql = "EXEC stp_upd_GeneraSecuenciaRCD_RCA '" & psFecha & "','" & psServConsol & "'," & IIf(pbRCDTransf, 1, 0) & ""
 Sql = "EXEC stp_upd_GeneraSecuenciaRCD_RCA '" & psFecha & "','" & psServConsol & "'," & pbRCDTransf & ""
 'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
 oCon.Ejecutar (Sql)
 oConec.CierraConexion
 Set oConec = Nothing
End Sub

'Public Function ValidaExisteRCD_RCA(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Boolean = False, Optional ByVal psTipoCond As String = "") As Boolean
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Public Function ValidaExisteRCD_RCA(ByVal psFecha As String, ByVal psServConsol As String, Optional ByVal pbRCDTransf As Integer, Optional ByVal psTipoCond As String = "") As Boolean
Dim oConec As COMConecta.DCOMConecta
Dim psSql As String
Dim pbValidaRCD_RCA As Boolean
Dim oRs As ADODB.Recordset
Set oConec = New COMConecta.DCOMConecta
Set oRs = New ADODB.Recordset
oConec.AbreConexion
'Valida si existe en la Consolidada RCD_RCDT_RCA
'psSql = "Exec stp_sel_ValidaExisteRCD_RCA '" & psFecha & "','" & psServConsol & "'," & IIf(pbRCDTransf, 1, 0) & ", '" & psTipoCond & "'"
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc

psSql = "Exec stp_sel_ValidaExisteRCD_RCA '" & psFecha & "','" & psServConsol & "'," & pbRCDTransf & ", '" & psTipoCond & "'"
Set oRs = oConec.CargaRecordSet(psSql)

If Not oRs.BOF And Not oRs.EOF Then
   pbValidaRCD_RCA = IIf(IsNull(oRs!bValidaRCD_RCA), True, oRs!bValidaRCD_RCA)
Else
   pbValidaRCD_RCA = True
End If
ValidaExisteRCD_RCA = pbValidaRCD_RCA
oConec.CierraConexion
Set oConec = Nothing
End Function
'***************END NAGL 202007 - ACTA N°046-2020*********************'
