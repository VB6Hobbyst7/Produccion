VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCredDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private vsCadImpresion As String
Private vsCadImpreAmpli As String 'CTI3 FERIMORO : ERS082-2018
Private lsNegritaOn As String
Private lsNegritaOff  As String
Private lsSaltoLin As String
Private lsTab As String
Private EspHorxPag As Integer
Private EspVerxPag As Integer
Private nPuntPag As Integer
Private sCadCab As String
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Public rsAnalista As ADODB.Recordset
Dim rsAnaProducto As ADODB.Recordset

Dim sTelefono As String
'Dim sConvenio As String
'Dim sDireccionTrabajo As String

Dim oImpre As COMFunciones.FCOMImpresion
Dim oCad As COMFunciones.FCOMCadenas
Dim oImpresora As COMFunciones.FCOMVarImpresion

Public Event ProgresoBarra(ByVal i As Long, ByVal lsTitulo As String, ByVal lsSubtitulo As String)
Public Event IniciaBarra(ByVal lnTotal As Long)
Public Event FinalizaBarra()

Public bAplicaComision As Boolean 'RIRO 20200706

'*** PEAC 20080412
Private Sub ImprimeCabeceraDocumentoCred(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "", Optional psCodana As String = "", Optional psNomCli As String = "", Optional psnumcta As String = "", _
                                    Optional pstipocredito As String = "", Optional psCodCli As String = "", Optional psdocnatural As String = "", Optional psdocjuri As String = "")

    Dim lsOperacion As String
    If Mid(psnumcta, 6, 3) = "515" Or Mid(psnumcta, 6, 3) = "516" Then
        lsOperacion = "OPERACION"
    Else
        lsOperacion = "CREDITO"
    End If
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(50) & "FECHA :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 45, 0) & Space(30) & "ANALISTA : " & psCodana & Chr$(10)
        psCadImp = psCadImp & psTab & "CLIENTE : " & oImpre.ImpreFormat(psNomCli, 35, 0) & Space(30) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & lsOperacion & ": " & oImpre.ImpreFormat(psnumcta, 35, 0) & Space(29) & pstipocredito & Chr$(10)
        psCadImp = psCadImp & psTab & "COD. CLIENTE : " & psCodCli & Space(15) & "DOC. NATURAL : " & psdocnatural & Space(10) & "DOC. JURIDICO : " & psdocjuri & Chr$(10)
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(50) & "FECHA :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 45, 0) & Space(30) & "ANALISTA : " & psCodana & Chr$(10)
        psCadImp = psCadImp & psTab & "CLIENTE : " & oImpre.ImpreFormat(psNomCli, 35, 0) & Space(30) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & lsOperacion & ": " & oImpre.ImpreFormat(psnumcta, 35, 0) & Space(29) & pstipocredito & Chr$(10)
        psCadImp = psCadImp & psTab & "COD. CLIENTE : " & psCodCli & Space(15) & "DOC. NATURAL : " & psdocnatural & Space(10) & "DOC. JURIDICO : " & psdocjuri & Chr$(10)
    End If
End Sub


Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
                                    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
                                    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
                                    Optional psSubTitulo As String = "") 'DAOR 20070816, Se aumentó subtitulo
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "FECHA :" & psFechaHora & Chr$(10)
        
        'by Capi 05032008
        'psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space(71) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    
    Else
        '*** PEAC 20080722
        'psCadImp = psCadImp & psTab & psNomCmac & Space(90) & "FECHA   :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & psNomCmac & Space(55) & "FECHA   :" & psFechaHora & Chr$(10)
        
        'By Capi 05032008
        'psCadImp = psCadImp & psTab & oImpre.ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        
        '*** PEAC 20080722
        'psCadImp = psCadImp & psTab & Space(114) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space(79) & "USUARIO : " & psCodUsu & Chr$(10)
        
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
        
    End If
End Sub

Public Function ImprimeSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, Optional ByVal psNomCmact As String = "") As String
    

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim R3 As ADODB.Recordset '*** PEAC 20091211
Dim sCadImp As String
Dim oCred As COMDCredito.DCOMCredito
Dim nCont As Integer
Dim nResto As Integer
Dim sCadImpTmp As String
Dim scargo1 As String
Dim sCadImpTmp2 As String

    On Error GoTo ErrorImprimeSolicitud
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosSolicitudCredito(psCtaCod)
    Set oDCred = Nothing
    
    Set oCred = New COMDCredito.DCOMCredito
    Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
    Set R3 = oCred.RecuperaRelacPers(psCtaCod) '*** PEAC 20091211
    Set oCred = Nothing
        
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, IIf(R2.RecordCount > 0, "SOLICITUD DE CREDITO REFINANCIADO", IIf(Mid(psCtaCod, 6, 3) = "515", "SOLICITUD DE ARRENDAMIENTO FINANCIERO", IIf(Mid(psCtaCod, 6, 3) = "516", "SOLICITUD DE ARRENDAMIENTO FINANCIERO", "SOLICITUD DE CREDITO"))), psNomCmact, lsTab, False)
    If Not R.BOF And Not R.EOF Then
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
            sCadImp = sCadImp & lsTab & lsNegritaOn & "Operación AF : " & psCtaCod & Space(15) & R!cTipoCred & " - " & R!cCondicion & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & Space(15) & R!cTipoCred & " - " & R!cCondicion & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(5) & oCad.PstaNombre(R!cPersNombre) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Doc. Natural   : " & oImpre.ImpreFormat(IIf(IsNull(R!DNI), "", R!DNI), 12, 0) & "Doc. Juridico    : " & oImpre.ImpreFormat(IIf(IsNull(R!RUC), "", R!RUC), 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Nro Hijos      : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(IIf(IsNull(R!Nacimiento), CDate(psFecSis), R!Nacimiento)), "#0") & " Años " & Space(20) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc & lnSaltoLinDoc
       ' sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE DE INGRESO" & Space(30) & IIf(IsNull(R!cFteInd), "DEPENDIENTE", "INDEPENDIENTE") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
'         sCadImp = sCadImp & lsTab & "FUENTE DE INGRESO" & Space(30) & IIf(CInt((R!nTipoFuente)) = 1, "DEPENDIENTE", "INDEPENDIENTE") & lnSaltoLinDoc & lnSaltoLinDoc

        '*** PEAC 20080515
'        sCadImp = sCadImp & lsTab & "Razon Social           : " & R!cRazonSocial & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "Direccion              : " & R!cDireccRazon & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "Actividad              : " & R!cCiiuDescripcion & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "Inicio de Actividades  : " & IIf(R!dpersnacCreac = 0, "", Format(R!dpersnacCreac, "dd/mm/yyyy")) & lnSaltoLinDoc
        
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
            sCadImp = sCadImp & lsTab & lsNegritaOn & "DATOS PRINCIPALES DEL ARRENDAMIENTO FINANCIERO" & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lsTab & lsNegritaOn & "DATOS PRINCIPALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & "Analista Responsable  : " & R!cAnalista & lnSaltoLinDoc
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
            sCadImp = sCadImp & lsTab & "Destino de  Operación : " & R!cDestino & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "Condicion de Operación: " & R!cCondicion & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lsTab & "Destino del Credito   : " & R!cDestino & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "Condicion del Credito : " & R!cCondicion & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & "Fecha de Asignacion   : " & Format(R!dFecAsig, "dd/mm/yyyy") & lnSaltoLinDoc
        If Not IsNull(R!cNomInst) Then
            sCadImp = sCadImp & lsTab & "Institucion           : " & R!cNomInst & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "Codigo Modular        : " & IIf(IsNull(R!cCodModular), "", R!cCodModular) & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Monto Solicitado      : " & Format(R!nMontoSol, "#0.00") & Space(15) & R!cMonedaSol & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cuotas Solicitadas    : " & Trim(Str(R!nCuotasSol)) & lnSaltoLinDoc
        'sCadImp = sCadImp & lsTab & "Plazo Solicitado      : " & Trim(Str(R!nPlazoSol)) & " Dia(s)" & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Plazo Solicitado      : " & Trim(Str(R!nPlazoSol)) & " Dia(s)" & lnSaltoLinDoc 'APRI20180215 ERS036-2017
        sCadImp = sCadImp & lsTab & "Envío de Estado de Cuenta: " & Trim(R!cEnvioEstadCta) & lsNegritaOff & lnSaltoLinDoc 'APRI20180215 ERS036-2017
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        
        'ARCV 29-01-2007
        'Set oCred = New COMDCredito.DCOMCredito
        'Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
        '--------
        If R2.RecordCount > 0 Then
            sCadImp = sCadImp & lsTab & "Creditos a Refinanciar :" & lnSaltoLinDoc & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Credito", 20) & oImpre.ImpreFormat("Capital", 10) & oImpre.ImpreFormat("Interes", 20) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
            Do While Not R2.EOF
                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(R2!cCtaCodref, 18) & oImpre.ImpreFormat(R2!nCapitalRef, 8) & oImpre.ImpreFormat(R2!nInteresRef, 8) & lnSaltoLinDoc
                R2.MoveNext
            Loop
            sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        End If
        R2.Close
        'Set oCred = Nothing
                
        'ARCV 01-02-2007 (Formato para las firmas)
        sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
        
        '*** PEAC 20091211
        nCont = 0
        If Not R3.BOF And Not R3.EOF Then

            Do While Not R3.EOF
            
                nCont = nCont + 1
                nResto = nCont Mod 2
                If nResto = 1 Then
                    sCadImpTmp = ""
                    sCadImpTmp2 = ""
                    scargo1 = ""
                    scargo1 = Trim(R3!cConsDescripcion)
                    sCadImpTmp = sCadImpTmp & lsTab & Space(5) & "____________________________" & lnSaltoLinDoc & lnSaltoLinDoc
                    sCadImpTmp = sCadImpTmp & lsTab & Space(5) & "      FIRMA " & scargo1 & lnSaltoLinDoc
                Else
                    sCadImpTmp = ""
                    sCadImpTmp2 = ""
                    sCadImpTmp2 = sCadImpTmp2 & lsTab & Space(5) & "____________________________" & Space(15) & "______________________________" & lnSaltoLinDoc & lnSaltoLinDoc
                    sCadImpTmp2 = sCadImpTmp2 & lsTab & Space(5) & "      FIRMA " & scargo1 & Space(Len(Space(32)) - Len(scargo1)) & "         FIRMA " & Trim(R3!cConsDescripcion) & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
                End If
                
                If nCont > 1 And nResto = 0 Then
                    sCadImp = sCadImp & sCadImpTmp2
                End If
                
            R3.MoveNext
            Loop
        
            If nCont > 1 And nResto = 1 Then
                sCadImp = sCadImp & sCadImpTmp
            Else 'sirve cuando las personas relacionadas solo es uno o pares (2,4,6,8,n...)
                sCadImp = sCadImp & sCadImpTmp
            End If

        End If
        
                    '*** PEAC 20080412
            '        sCadImp = sCadImp & lsTab & Space(5) & "____________________________" & lnSaltoLinDoc & lnSaltoLinDoc
            '        sCadImp = sCadImp & lsTab & Space(5) & "      Firma Solicitante     " & lnSaltoLinDoc
            '
            ''        sCadImp = sCadImp & lsTab & Space(5) & "____________________________" & Space(15) & "______________________________" & lnSaltoLinDoc & lnSaltoLinDoc
            ''        sCadImp = sCadImp & lsTab & Space(5) & "      Firma Solicitante     " & Space(15) & "         Firma Cónyuge        " & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
            ''        sCadImp = sCadImp & lsTab & Space(5) & "____________________________" & Space(15) & "______________________________" & lnSaltoLinDoc & lnSaltoLinDoc
            ''        sCadImp = sCadImp & lsTab & Space(5) & "           Firma Aval       " & Space(15) & "           Firma Aval         " & lnSaltoLinDoc
                    '*** FIN PEAC
                    '------------
        
        '*** FIN PEAC

        sCadImp = sCadImp & Chr$(12)
    End If
    R.Close
    Set R = Nothing
    ImprimeSolicitud = sCadImp
    Exit Function

ErrorImprimeSolicitud:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Private Function lnSaltoLinDoc() As String
    nPuntPag = nPuntPag + 1
    If nPuntPag > EspVerxPag Then
        nPuntPag = 0
        lnSaltoLinDoc = Chr$(12) & sCadCab
    Else
        lnSaltoLinDoc = Chr$(10)
    End If
    
End Function

'Se volvio Funcion para pasar al Cliente la CADENA

Public Function ImprimeBoleta(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pstipocredito As String, ByVal pnCapPagado As Double, _
            ByVal pnIntCompPagado As Double, ByVal pnIntCompVencPagado As Double, ByVal pnIntMorPagado As Double, ByVal pnGastoPagado As Double, _
            ByVal pnIntGraciaPagado As Double, ByVal pnIntSuspReprog As Double, ByVal pnSaldoCap As Double, _
            ByVal psProxFecPago As String, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", _
            Optional ByVal pbPagoCheque As Boolean = False, Optional ByVal psNumCheque As String = "", Optional ByVal psCodCMAC As String = "", _
            Optional ByVal pnITF As Double = 0#, Optional ByVal pnInteresDesagio As Double = 0, Optional pbOpeRecCMAC As Boolean = False, Optional pbCargoAuto As Boolean = False, _
            Optional ByVal pbPagoDacion As Boolean = False, Optional ByVal pbImpTMU As Boolean = False, Optional ByVal pbMantPrepago As Boolean = False, _
            Optional ByVal pnModalMantPrePago As Integer = -1, Optional ByVal psDOI As String = "", Optional ByVal pnMovNro As Long = 0, _
            Optional ByVal pConnBase As COMDCredito.DCOMCredActBD = Nothing, Optional ByVal pnMotivoCanc As Integer = 0, Optional ByVal pnMontoComPreCanc As Double = 0, _
            Optional ByVal pnTipoPagoCred As CredTiposPagosNormales = 0, Optional ByVal pnSaldoCapPend As Double = 0, Optional ByVal pnSaldoCapPorVenc As Double = 0, _
            Optional ByVal pnTipoAdelantoCuota As Integer = 0, Optional ByVal pnMotivoCancela As Variant = Nothing) As String
            'JUEZ 20121001 Se le agregó el tipo de dato a pnIntCompVencPagado
            'JUEZ 20121226 Se agregó pbMantPrepago, pnModalMantPrePago y psDOI
            'AMDO 20130320 Se agrego pnMovNro
            'AMDO 20130326 Se agrego pConnBase
            'AMDO 20130326 Se agregó pnMotivoCanc
            'JUEZ 20130925 Se agregó parametro pnMontoComPreCanc
            'JUEZ 20141212 Se agregé pnTipoPagoCred,pnSaldoCapPend,pnSaldoCapPorVenc
            'JUEZ 20150415 Se agregó pnTipoAdelantoCuota
'Dim nFicSal As Integer
            'CTI3 ERS082-2018 SE AGREGO pnMotivoCancela
'TypeName(pnMovNroRVD) = "Nothing"
'JOEP20211102 Mejora en pago anticipado
    Dim oTpCalend As COMDCredito.DCOMCredito
    Dim nTpCalendIteTra As Integer
    nTpCalendIteTra = 0
    Set oTpCalend = New COMDCredito.DCOMCredito
    nTpCalendIteTra = oTpCalend.GetTipoCalendIterativoTradi(psCtaCod)
    Set oTpCalend = Nothing
'JOEP20211102 Mejora en pago anticipado

    vsCadImpresion = ""
'    nFicSal = FreeFile

'    Open sLpt For Output As nFicSal
    
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    'Print #nFicSal, ""
    
    'JUEZ 20131021 *******************************************************************************
    Dim oGen As New COMDConstSistema.DCOMGeneral, nEstadoCampDatos As Integer
    nEstadoCampDatos = oGen.LeeConstSistema(gConstSistEstadoCampActualizaDatos)
    If nEstadoCampDatos = 1 Then
        If pbImpTMU = False Then
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn
            vsCadImpresion = vsCadImpresion & "                            * *         " & Space(11) & "                            * *         " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "                         *      *       " & Space(11) & "                         *      *       " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "                        *               " & Space(11) & "                        *               " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "      * *              *          *     " & Space(11) & "      * *              *          *     " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "    *     *      *   *             *    " & Space(11) & "    *     *      *   *             *    " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "   *        *  *   *                    " & Space(11) & "   *        *  *   *                    " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "  *           *                     *   " & Space(11) & "  *           *                     *   " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & " *                                   *  " & Space(11) & " *                                   *  " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "*                                     * " & Space(11) & "*                                     * " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & Space(11) & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & lsNegritaOff & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn & "C"
            vsCadImpresion = vsCadImpresion & "                            * *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "                         *      *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "                        *         " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "      * *              *          *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "    *     *      *   *             *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "   *        *  *   *                " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "  *           *                     *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & " *                                   *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "*                                     *" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & lsNegritaOff & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin
        End If
    End If
    'END JUEZ ************************************************************************************
    
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, psCuotasPagadas, _
            psFecha, psHora, psPersNombre, psNroTransac, psCtaCod, pstipocredito, psNomCmacAbrev, psCodCMAC, pbImpTMU, , pnMotivoCancela) 'CTI3 ERS082-2018: pnMotivoCancela
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    
    'AMDO 20130326 ***************************************
        Dim oVarSist As COMDConstSistema.DCOMGeneral
        Set oVarSist = New COMDConstSistema.DCOMGeneral
        Dim pCrediPremiazo As Integer
        pCrediPremiazo = oVarSist.LeeConstSistema(432)
        Set oVarSist = Nothing
    
        Dim oCredito As COMNCredito.NCOMCredito
        Dim puntosGanados As Integer
        Dim puntoTotales As Integer
    'END AMDO ********************************************
    
    '45 menos la longitud de su etiqueta
If psCodCMAC = "102" Then
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & "Cheque  :" & Space(2) & oImpre.ImpreFormat(psNumCheque, 12) & lsSaltoLin
    End If
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
    Else
        vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & lsSaltoLin
    End If
    
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(16) & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & oImpre.ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    
'   Print #nFicSal, vsCadImpresion
'   Print #nFicSal, ""
'   Close #nFicSal
    ImprimeBoleta = vsCadImpresion
Else
    If pbPagoCheque Then
        If pbImpTMU = False Then
            vsCadImpresion = vsCadImpresion & "Cheque  :" & Space(2) & oImpre.ImpreFormat(psNumCheque, 12)
            vsCadImpresion = vsCadImpresion & Space(26) & "Cheque  :" & Space(2) & oImpre.ImpreFormat(psNumCheque, 12) & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & "Cheque  :" & Space(2) & oImpre.ImpreFormat(psNumCheque, 12) & lsSaltoLin
        End If
    End If
    If pbPagoCheque Then
        If pbImpTMU = False Then
            If pbOpeRecCMAC = True Then
                vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CHEQUE RECEP. " & String(10, "-")
                vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & "PAGO CHEQUE RECEP. " & String(10, "-") & lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-")
                vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
            End If
        Else
            If pbOpeRecCMAC = True Then
                vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CHEQUE RECEP. " & String(10, "-") & lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
            End If
        End If
    Else
        If pbImpTMU = False Then
            If pbOpeRecCMAC = True Then
                vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO.EFEC.RECEPC. " & String(10, "-")
                vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO.EFEC.RECEPC. " & String(10, "-") & lsSaltoLin
            Else
                If pbCargoAuto Then
                    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO CARGO CUENTA " & String(10, "-")
                    vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO CARGO CUENTA " & String(10, "-") & lsSaltoLin
                Else
                    If pbPagoDacion Then
                        vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO CON DACION " & String(10, "-")
                        vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO CON DACION " & String(10, "-") & lsSaltoLin
                    Else
                    'CTI3: ERS082-2018
                    '    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-")
                    '    vsCadImpresion = vsCadImpresion & Space(12) & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-") & lsSaltoLin
                        If IsArray(pnMotivoCancela) Then 'JOEP201190411 Mejora
                            If pnMotivoCancela(0) = "CANCELADO" And pnMotivoCancela(1) = True Then
                                vsCadImpresion = vsCadImpresion & String(8, "-") & " CANCELACION CREDITO  " & String(9, "-")
                                vsCadImpresion = vsCadImpresion & Space(12) & String(8, "-") & " CANCELACION CREDITO  " & String(9, "-") & lsSaltoLin
                            Else
                                If pnMotivoCancela(0) = "CANCELADO" And pnMotivoCancela(1) = False Then
                                    vsCadImpresion = vsCadImpresion & String(8, "-") & " CANCELACION CREDITO  " & String(9, "-")
                                    vsCadImpresion = vsCadImpresion & Space(12) & String(8, "-") & " CANCELACION CREDITO  " & String(9, "-") & lsSaltoLin
                                Else
                                    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-")
                                    vsCadImpresion = vsCadImpresion & Space(12) & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-") & lsSaltoLin ' & Chr(13)
                                End If
                            End If
                        'JOEP201190411 Mejora
                        Else
                            vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-")
                            vsCadImpresion = vsCadImpresion & Space(12) & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-") & lsSaltoLin ' & Chr(13)
                        End If
                        'JOEP201190411 Mejora
                    End If
                End If
            End If
        Else
            If pbOpeRecCMAC = True Then
                vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO.EFEC.RECEPC. " & String(10, "-") & lsSaltoLin
            Else
                If pbCargoAuto Then
                    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO CARGO CUENTA " & String(10, "-") & lsSaltoLin
                Else
                    If pbPagoDacion Then
                        vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO CON DACION " & String(10, "-") & lsSaltoLin
                    Else
                        vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-") & lsSaltoLin
                    End If
                End If
            End If
        End If
    End If
    
    If pbImpTMU = False Then
        vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        
        If pnInteresDesagio > 0 Then
            vsCadImpresion = vsCadImpresion & "Interes Desagio (-)  :" & oImpre.ImpreFormat(pnInteresDesagio, 12)
            vsCadImpresion = vsCadImpresion & Space(14) & "Interes Desagio (-)   :" & oImpre.ImpreFormat(pnInteresDesagio, 12) & lsSaltoLin
        Else
        'JOEP20211102 Mejora en pago anticipado
            If nTpCalendIteTra = 1 Then
                vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado + pnIntGraciaPagado, 12)
                vsCadImpresion = vsCadImpresion & Space(14) & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado + pnIntGraciaPagado, 12) & lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12)
                vsCadImpresion = vsCadImpresion & Space(14) & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
            End If
        'JOEP20211102 Mejora en pago anticipado
            'vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12)
            'vsCadImpresion = vsCadImpresion & Space(14) & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
        End If
    Else
        vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        If pnInteresDesagio > 0 Then
            vsCadImpresion = vsCadImpresion & "Interes Desagio (-)  :" & oImpre.ImpreFormat(pnInteresDesagio, 12) & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & oImpre.ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
        End If
    End If
    
    If pbImpTMU = False Then
        'vsCadImpresion = vsCadImpresion & "Mora/Penalidad :" & Space(6) & oImpre.ImpreFormat(pnIntMorPagado, 12)
        vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPagado, 12) 'ORCR20140714
        'vsCadImpresion = vsCadImpresion & Space(14) & "Mora/Penalidad :" & Space(6) & oImpre.ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Space(14) & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin 'ORCR20140714
        vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastoPagado, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    
    'JOEP20211102 Mejora en pago anticipado
        If nTpCalendIteTra = 1 Then
        Else
            vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12)
            vsCadImpresion = vsCadImpresion & Space(14) & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
        End If
    'JOEP20211102 Mejora en pago anticipado
    
        'vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(14) & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & oImpre.ImpreFormat(pnIntSuspReprog, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "Int. Susp y Reprog:" & Space(3) & oImpre.ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "I.T.F.            :" & Space(3) & oImpre.ImpreFormat(pnITF, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "I.T.F.            :" & Space(3) & oImpre.ImpreFormat(pnITF, 12) & lsSaltoLin
        If pnMontoComPreCanc > 0 Then 'JUEZ 20130925
            vsCadImpresion = vsCadImpresion & "Comision PreCancelacion:" & oImpre.ImpreFormat(pnMontoComPreCanc, 10)
            vsCadImpresion = vsCadImpresion & Space(14) & "Comision PreCancelacion:" & oImpre.ImpreFormat(pnMontoComPreCanc, 10) & lsSaltoLin
        End If
        vsCadImpresion = vsCadImpresion & String(38, "-")
        vsCadImpresion = vsCadImpresion & Space(13) & String(39, "-") & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF - pnInteresDesagio, 12)
        vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF + pnMontoComPreCanc - pnInteresDesagio, 12) 'JUEZ 20130925
        'vsCadImpresion = vsCadImpresion & Space(14) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF - pnInteresDesagio, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Space(14) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF + pnMontoComPreCanc - pnInteresDesagio, 12) & lsSaltoLin 'JUEZ 20130925
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        'JUEZ 201421212 *********************************************************
        If pnTipoPagoCred = gPagoVencSaldoPend Then
            vsCadImpresion = vsCadImpresion & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8)
            vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoVencSaldoPorVenc Then
            vsCadImpresion = vsCadImpresion & "Saldo de Cuota por Vencer:" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8)
            vsCadImpresion = vsCadImpresion & Space(14) & "Saldo de Cuota por Vencer :" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoNorEnFechaCuotaSaldoPend Then
            vsCadImpresion = vsCadImpresion & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8)
            vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoNorAntesFechaCuotaSaldoPorVenc Then
            vsCadImpresion = vsCadImpresion & "Saldo de Cuota por Vencer:" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8)
            vsCadImpresion = vsCadImpresion & Space(14) & "Saldo de Cuota por Vencer :" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8) & lsSaltoLin
        End If
        'END JUEZ ***************************************************************
        'COMENTADO POR CTI3
        If IsArray(pnMotivoCancela) Then  'JOEP201190411 Mejora
            If pnMotivoCancela(0) = "CANCELADO" Then
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("USUARIO : " & psProxFecPago & Space(23) & psCodUsu, 51, 0) & oImpre.ImpreFormat("USUARIO : " & psProxFecPago & Space(23) & psCodUsu, 53, 0) & Chr(13) '& lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12)
                vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin '& Chr(13)
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(4) & psCodUsu, 51, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(4) & psCodUsu, 53, 0) & Chr(13) '& lsSaltoLin
            End If
        'JOEP201190411 Mejora
        Else
            vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12)
            vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin '& Chr(13)
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(4) & psCodUsu, 51, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(4) & psCodUsu, 53, 0) & Chr(13) '& lsSaltoLin
        End If
        'JOEP201190411 Mejora
        'vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12)
        'vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 51, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
        'END CTI3
        'AMDO 20130320 CREDIPREMIAZO ********************
        If pCrediPremiazo = 1 Then
            If pnMotivoCanc = 0 Then ' Si no es cancelacion anticipada
            Set oCredito = New COMNCredito.NCOMCredito
                If Mid(psCtaCod, 6, 3) = "517" And Not pConnBase Is Nothing Then
                    puntosGanados = pConnBase.DevuelveCantPuntosMovCredito(pnMovNro, psCtaCod)
                    puntoTotales = pConnBase.DevuelveCantPuntosTotalCredito(psCtaCod)
                Else
                    puntosGanados = oCredito.DevuelveCantPuntosMovCredito(pnMovNro, psCtaCod)
                    puntoTotales = oCredito.DevuelveCantPuntosTotalCredito(psCtaCod)
                End If
            Set oCredito = Nothing
        
                If puntosGanados > 0 Then
                    '' AMDO20130607 Se adecuo el voucher del CrediPremiazo segun el TI-ERS075-2013
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "******** F E L I C I T A C I O N E S  ********" & Space(6) & "******** F E L I C I T A C I O N E S  ********"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "**" & Space(42) & "**" & Space(6) & "**" & Space(42) & "**"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "** Usted ha ganado " & puntosGanados & " opciones, Siga pagando **" & Space(6) & "** Usted ha ganado " & puntosGanados & " opciones, Siga pagando **"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "** puntualmente sus cuotas y ganarás más    **" & Space(6) & "** puntualmente sus cuotas y ganarás más    **"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "** opciones para participar en los sorteos  **" & Space(6) & "** opciones para participar en los sorteos  **"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "** de CrediPremiazo." & Space(24) & "**" & Space(6) & "** de CrediPremiazo." & Space(24) & "**"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "**" & Space(42) & "**" & Space(6) & "**" & Space(42) & "**"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "** Opciones acumuladas a la fecha: " & oImpre.Justifica(CStr(puntoTotales), 9) & "**" & Space(6) & "** Opciones acumuladas a la fecha: " & oImpre.Justifica(CStr(puntoTotales), 9) & "**"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "**********************************************" & Space(6) & "**********************************************"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
                End If
            End If
        End If
        'END AMDO  *************************************
        'JUEZ 20121226 *********************************************************
        If pbMantPrepago Then
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn & "El cliente deja  constancia expresa  de" & Space(12) & "El cliente deja  constancia expresa  de" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "que  al  haber  realizado   este   pago" & Space(12) & "que  al  haber  realizado   este   pago" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "anticipado,   ha   escogido   que    su" & Space(12) & "anticipado,   ha   escogido   que    su" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "calendario sea reestructurado  bajo  la" & Space(12) & "calendario sea reestructurado  bajo  la" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "siguiente modalidad:" & Space(31) & "siguiente modalidad: " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "* Reducción del " & IIf(pnModalMantPrePago = 0, "Número de Cuotas" & Space(19), "Monto de las Cuotas" & Space(16))
            vsCadImpresion = vsCadImpresion & "* Reducción del " & IIf(pnModalMantPrePago = 0, "Número de Cuotas", "Monto de las Cuotas") & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & Space(31) & "____________________" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 11, 10)) & "DOI: " & psDOI & Space(IIf(Len(psDOI) = 8, 38, 35)) & "DOI: " & psDOI & lsNegritaOff
        'END JUEZ **************************************************************
        'JUEZ 20150415 *****************************************************************************************
        'JUEZ 20141212 *********************************************************
        'Else
        '   If pnTipoPagoCred = gPagoNorMaxDosCuotasPorVenc Then
        '       vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn & "El cliente deja  constancia expresa  de" & Space(12) & "El cliente deja  constancia expresa  de" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & "que  al  haber  realizado   este   pago" & Space(12) & "que  al  haber  realizado   este   pago" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & "ha  escogido  realizar  un  pago  en la" & Space(12) & "ha  escogido  realizar  un  pago  en la" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & "modalidad 'Adelanto de Cuota'          " & Space(12) & "modalidad 'Adelanto de Cuota'          " & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & Space(31) & "____________________" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 11, 10)) & "DOI: " & psDOI & Space(IIf(Len(psDOI) = 8, 38, 35)) & "DOI: " & psDOI & lsNegritaOff
        '   End If
        ElseIf pnTipoAdelantoCuota = 1 Then
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn & "El   tipo   de    pago   realizado   es" & Space(12) & "El   tipo   de    pago   realizado   es" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "'Adelanto de Cuota'" & Space(32) & "'Adelanto de Cuota'" & lsNegritaOff
        ElseIf pnTipoAdelantoCuota = 2 Then
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn & "El cliente deja  constancia expresa  de" & Space(12) & "El cliente deja  constancia expresa  de" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "que  al  haber  realizado   este   pago" & Space(12) & "que  al  haber  realizado   este   pago" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "ha  escogido  realizar  un  pago  en la" & Space(12) & "ha  escogido  realizar  un  pago  en la" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "modalidad 'Adelanto de Cuota'          " & Space(12) & "modalidad 'Adelanto de Cuota'          " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & Space(31) & "____________________" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 11, 10)) & "DOI: " & psDOI & Space(IIf(Len(psDOI) = 8, 38, 35)) & "DOI: " & psDOI & lsNegritaOff
        End If
        'END JUEZ **************************************************************
        'END JUEZ **********************************************************************************************
         vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
    Else
        'vsCadImpresion = vsCadImpresion & "Mora/Penalidad :" & Space(6) & oImpre.ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & oImpre.ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin 'ORCR20140714
        vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & oImpre.ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & oImpre.ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & oImpre.ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "I.T.F.            :" & Space(3) & oImpre.ImpreFormat(pnITF, 12) & lsSaltoLin
        If pnMontoComPreCanc > 0 Then 'JUEZ 20130925
            vsCadImpresion = vsCadImpresion & "Comision PreCancelacion:" & oImpre.ImpreFormat(pnMontoComPreCanc, 10) & lsSaltoLin
        End If
        vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF - pnInteresDesagio, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF + pnMontoComPreCanc - pnInteresDesagio, 12) & lsSaltoLin 'JUEZ 20130925
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        'JUEZ 201421212 *********************************************************
        If pnTipoPagoCred = gPagoVencSaldoPend Then
            vsCadImpresion = vsCadImpresion & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoVencSaldoPorVenc Then
            vsCadImpresion = vsCadImpresion & "Saldo de Cuota por Vencer:" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoNorEnFechaCuotaSaldoPend Then
            vsCadImpresion = vsCadImpresion & "Saldo Pendiente de Cuota :" & oImpre.ImpreFormat(pnSaldoCapPend, 8) & lsSaltoLin
        ElseIf pnTipoPagoCred = gPagoNorAntesFechaCuotaSaldoPorVenc Then
            vsCadImpresion = vsCadImpresion & "Saldo de Cuota por Vencer:" & oImpre.ImpreFormat(pnSaldoCapPorVenc, 8) & lsSaltoLin
        End If
        'END JUEZ ***************************************************************
        vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 51, 0) & lsSaltoLin
        'AMDO 20130320 CREDIPREMIAZO ********************
        If pCrediPremiazo = 1 Then
            If pnMotivoCanc = 0 Then 'Si no es cancelacion anticipada
                Set oCredito = New COMNCredito.NCOMCredito
                    If Mid(psCtaCod, 6, 3) = "517" And Not pConnBase Is Nothing Then
                        puntosGanados = pConnBase.DevuelveCantPuntosMovCredito(pnMovNro, psCtaCod)
                        puntoTotales = pConnBase.DevuelveCantPuntosTotalCredito(psCtaCod)
                    Else
                        puntosGanados = oCredito.DevuelveCantPuntosMovCredito(pnMovNro, psCtaCod)
                        puntoTotales = oCredito.DevuelveCantPuntosTotalCredito(psCtaCod)
                    End If
                Set oCredito = Nothing
    
                If puntosGanados > 0 Then
                    '' AMDO20130607 Se adecuo el voucher del CrediPremiazo segun el TI-ERS075-2013
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "***** F E L I C I T A C I O N E S  *****"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "*" & Space(38) & "*"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "* Usted ha ganado " & puntosGanados & " opciones, Siga     *" & lsSaltoLin & "* pagando puntualmente sus cuotas y    *"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "* ganarás más opciones para participar *" & lsSaltoLin & "* en los sorteos de CrediPremiazo." & Space(5) & "*"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "*" & Space(38) & "*"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "* Opciones acumuladas a la fecha: " & oImpre.Justifica(CStr(puntoTotales), 5) & "*"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & "****************************************"
                    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
                End If
            End If
        End If
        'END AMDO  *************************************
        'JUEZ 20121226 *********************************************************
        If pbMantPrepago Then
            vsCadImpresion = vsCadImpresion & lsSaltoLin & "El cliente deja  constancia expresa  de" & lsSaltoLin & "que  al  haber  realizado   este   pago" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "anticipado,   ha   escogido   que    su" & lsSaltoLin & "calendario sea reestructurado  bajo  la" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "siguiente modalidad: " & lsSaltoLin & "* Reducción del " & IIf(pnModalMantPrePago = 0, "Número de Cuotas", "Monto de las Cuotas") & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 12, 10)) & "DOI: " & psDOI & lsSaltoLin
        'END JUEZ **************************************************************
        'JUEZ 20150415 *****************************************************************************************
        'JUEZ 20141212 *********************************************************
        'Else
        '   If pnTipoPagoCred = gPagoNorMaxDosCuotasPorVenc Then
        '       vsCadImpresion = vsCadImpresion & lsSaltoLin & "El cliente deja  constancia expresa  de" & lsSaltoLin & "que  al  haber  realizado   este   pago" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & "ha  escogido  realizar  un  pago  en la" & lsSaltoLin & "modalidad 'Adelanto de Cuota'          " & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & lsSaltoLin
        '       vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 12, 10)) & "DOI: " & psDOI & lsSaltoLin
        '   End If
        ElseIf pnTipoAdelantoCuota = 1 Then
            vsCadImpresion = vsCadImpresion & lsNegritaOn & "EEl   tipo   de    pago   realizado   es" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "'Adelanto de Cuota'" & lsNegritaOff & lsSaltoLin
        ElseIf pnTipoAdelantoCuota = 2 Then
            vsCadImpresion = vsCadImpresion & lsNegritaOn & "EEl cliente deja  constancia expresa  de" & lsSaltoLin & "que  al  haber  realizado   este   pago" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & "ha  escogido  realizar  un  pago  en la" & lsSaltoLin & "modalidad 'Adelanto de Cuota'          " & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Space(8) & "____________________" & lsSaltoLin
            vsCadImpresion = vsCadImpresion & Space(IIf(Len(psDOI) = 8, 12, 10)) & "DOI: " & psDOI & lsNegritaOff & lsSaltoLin
        End If
        'END JUEZ **************************************************************
        'END JUEZ **********************************************************************************************
        vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Chr(27) & "m" & Chr(10)
    End If
        
    'LARI 20210201 MENSAJE DE AMPLIACIÓN********************************************
    'strTest = Mid$("Visual Basic", 3, 3) 'strTest would be "ual"
    Dim Mensaje As String '41 caracteres por linea
    Dim Mensaje1 As String
    Dim CantDig As Integer
    
    Mensaje = "Felicidades ya puedes ampliar tu crédito y tu ya vez que haces con eso, piensalo"
    
    Dim oCredito1 As COMNCredito.NCOMCredito
    Set oCredito1 = New COMNCredito.NCOMCredito
    

    Mensaje = oCredito1.ValidaAmpliacion(psCtaCod)
    If (Len(Mensaje) > 41) Then
         Do While (Len(Mensaje) > 41)
         Mensaje1 = Mid(Mensaje, 1, 41) & Space(9) & Mid(Mensaje, 1, 41)
         vsCadImpresion = vsCadImpresion & lsNegritaOn & Mensaje1 & lsSaltoLin
         Mensaje = Mid(Mensaje, 41, Len(Mensaje) - 40)
         
         If (Len(Mensaje) <= 41) Then
         CantDig = 41 - Len(Mensaje)
         Mensaje1 = Mensaje & Space(CantDig) & Space(9) & Mensaje
         vsCadImpresion = vsCadImpresion & lsNegritaOn & Mensaje1 & lsSaltoLin
         End If
         
         Loop
    Else
        Mensaje = Mensaje & Space(9) & Mensaje
        vsCadImpresion = vsCadImpresion & lsNegritaOn & Mensaje & lsSaltoLin
    End If
    
    '*******************************************************************************

    ImprimeBoleta = vsCadImpresion
End If
End Function

Private Sub ImprimeCabeceraBoletaPago(ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, ByVal psNroTransac As String, ByVal psCtaCod As String, _
        ByVal pstipocredito As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, Optional ByVal pbImpTMU As Boolean = False, _
        Optional ByVal pbPagoHonra As Boolean = False, Optional ByVal pnCancelado As Variant = Nothing)
       ' psNomAge = "Sede Institucional"
       'WIOR 20131230 AGREGO pbPagoHonra
        On Error GoTo ErrorImprimeCabeceraBoletaPago
        Dim lsTipoProducto As String
        Dim lsTipoProducto2 As String
        Dim J, tota As Integer 'CTI3
        Dim ultPago As String 'CTI3
        Dim a As Integer
        ultPago = ""
        lsTipoProducto = ""
        lsTipoProducto2 = ""
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
            lsTipoProducto = "ARRENDAMIENTO"
            lsTipoProducto2 = "Operacion"
        Else
            lsTipoProducto = "CREDITOS"
            lsTipoProducto2 = "Credito"
        End If
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        If psCodCMAC = "102" Then
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - " & lsTipoProducto & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Else
            If pbImpTMU = False Then
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - " & lsTipoProducto & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 50, 0) & Space(1) & Trim(oImpre.ImpreFormat(psNomCmacAbrev & " - " & lsTipoProducto & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
                    If IsArray(pnCancelado) Then 'JOEP201190411 Mejora
                        If pnCancelado(0) = "CANCELADO" Then
                            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 50, 0) & Space(1) & Trim(oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin 'CTI3
                        End If
                    End If
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & lsNegritaOff
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & "Cliente : " & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 41, 0) & Trim("Cliente : " & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin 'CTI3
                'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 51, 0) & Trim(oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin  'CTI3
                If Not pbPagoHonra Then 'WIOR 20131230
                
                     'CTI3 ERS 082-2018
                    If IsArray(pnCancelado) Then  'JOEP201190411 Mejora
                         If pnCancelado(0) = "CANCELADO" And pnCancelado(1) = True Then
                         
                           tota = Len(psCuotasPagadas)
                           For J = tota To 1 Step -1
                              If Mid(psCuotasPagadas, J, 1) <> "," Then ultPago = Mid(psCuotasPagadas, J, 1) + ultPago Else: Exit For
                           Next J
                         
                           vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(ultPago, 1, 15), 51, 0) & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(ultPago, 1, 15), 53, 0) & lsSaltoLin
                        
                         Else
                           If pnCancelado(0) = "CANCELADO" And pnCancelado(1) = False Then
                               vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
                           Else
                               vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
                           End If
                         End If
                    End If

                'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin 'CTI3
                'WIOR 20131230 *******************************
                Else
                    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(22), 51, 0) & oImpre.ImpreFormat(lsTipoProducto2 & " : " & psCtaCod & Space(22), 53, 0) & lsSaltoLin
                End If
                'WIOR FIN ***********************************9
            Else
                vsCadImpresion = vsCadImpresion & "C" & psNomCmacAbrev & "- " & lsTipoProducto & Space(2) & pstipocredito & lsSaltoLin
                vsCadImpresion = vsCadImpresion & psNomAge & " - " & psMoneda & Space(3) & Trim(psNroTransac) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & lsNegritaOff
                vsCadImpresion = vsCadImpresion & "Fecha : " & psFecha & Space(5) & "Hora : " & psHora & lsSaltoLin
                vsCadImpresion = vsCadImpresion & "Clte.: " & psPersNombre & lsSaltoLin
                If Not pbPagoHonra Then 'WIOR 20131230
                    vsCadImpresion = vsCadImpresion & lsTipoProducto2 & " : " & psCtaCod & Space(2) & "Cuota:" & psCuotasPagadas & lsSaltoLin
                'WIOR 20131230 *******************************
                Else
                    vsCadImpresion = vsCadImpresion & lsTipoProducto2 & " : " & psCtaCod & Space(2) & lsSaltoLin
                End If
                'WIOR FIN ***********************************9
            End If
        End If
        
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub

'*******************************************************************************************************
'Devuelve Un a Arreglo de Strings donde devuelve todos los documentos que se emiten en un desembolso
'Tanto en Abono a cuenta como en efectivo, el valor de 1 Indica papel boleta y el valor 2 indica papel carta
'*******************************************************************************************************

Public Function ImprimeDocumentosDesembolso(ByVal psCtaCod As String, ByVal pnMovNro As Long, _
    ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pdFecSis As Date, _
    Optional ByVal pbApeAboCta As Integer, Optional ByVal psctaAho As String = "", _
    Optional pnAhoMontoApe As Double = 0#, _
    Optional pnIntGanadoAbono As Double = 0#, Optional pnIntGanadoRetGastos As Double = 0#, _
    Optional pnIntGanadoRetCancel As Double = 0#, _
    Optional pnMontoAbono As Double = 0#, Optional pnMontoRetGastos As Double = 0#, _
    Optional pnMontoRetCancel As Double = 0#, _
    Optional pnSaldoDispAhoAbono As Double = 0#, Optional ByVal pnSaldoCntAhoAbono As Double = 0#, _
    Optional pnSaldoDispAhoRetGastos As Double = 0#, Optional pnSaldoCntAhoRetGastos As Double = 0#, _
    Optional pnSaldoDispAhoRetCancel As Double = 0#, Optional pnSaldoCntAhoRetCancel As Double = 0#, _
    Optional ByVal psNomCmac As String = "", Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, Optional pnTIF As Double = 0, _
    Optional ByVal pnMontoPoliza As Double = 0, Optional ByVal pMatDatosAhoEmp As Variant = Nothing, Optional ByVal psBolTransfGarInfoGas As String = "", _
     Optional ByVal pnMontoRetiroEfectivo As Double = 0#, Optional ByVal pnSaldoDispAhoRetEfectivo As Double = 0#, _
    Optional ByVal pnSaldoCntAhoRetEfectivo As Double = 0#, Optional ByVal pnIntGanadoRetEfectivo As Double = 0#, Optional ByVal pnRefina As Integer = -1, _
    Optional ByVal pnAmpliado As Integer, Optional ByRef pMatCuentaGastoCierre As Variant = "") As Variant
    'FRHU 20140228 RQ14006 Se agrego: pnMontoRetiroEfectivo,pnSaldoDispAhoRetEfectivo, pnSaldoCntAhoRetEfectivo y pnIntGanadoRetEfectivo
    'LUCV20160621, ERS004-2016. Agregó: pnRefina
    
Dim MatDocs() As String
Dim MatTempDesem(1, 2) As String
Dim MatTempDesemGastos(1, 2) As String
Dim MatTempDesemCancel() As String
Dim MatTempComprob(1, 2) As String
Dim MatTempPlan(2, 2) As String
Dim MatTempAhoApe(1, 2) As String
Dim MatTempAhoRetGastos(1, 2) As String
Dim MatTempAhoRetCancel(1, 2) As String
Dim MatTempAhoRetEfectivo(1, 2) As String 'FRHU20140228 RQ14006
Dim lsDescripcionBoleta1, lsDescripcionBoleta2 As String
Dim nTotReg As Integer
Dim oDCredDoc As COMDCredito.DCOMCredDoc
Dim oDCredito As COMDCredito.DCOMCredito
Dim oNCred As NCOMCredito 'BRGO 20111021
Dim R As ADODB.Recordset
Dim sCadTmp As String
Dim AcumTmp As Double
Dim i As Integer
Dim nCapital As Double
Dim nGasto As Double
Dim nNumTran As Integer
Dim nLin As Integer
Dim nMontoDesemb As Double
Dim nCalMiVivi As Boolean
Dim dProxFecPag As Date
Dim lsTItular As String 'ALPA 20121226
Dim sCadAmplTmp As String 'CTI3
Dim sCadDatosTmp(20) As String 'CTI3
'JUEZ 20131021 **************************************************
Dim oGen As New COMDConstSistema.DCOMGeneral, nEstadoCampDatos As Integer
nEstadoCampDatos = oGen.LeeConstSistema(gConstSistEstadoCampActualizaDatos)
'END JUEZ *******************************************************
'CTI3
Dim nCancelado(1) As Variant
nCancelado(0) = "CANCELADO"
nCancelado(1) = True
Dim mCalc As Double
Dim sEspa As Integer
Dim lsBoletaApeFMV As String

sEspa = 33

    nTotReg = 0
    '*****************************************************************************
    '****   Recuperar Boletas de Desembolso
    '*****************************************************************************
    
    Set oDCredito = New COMDCredito.DCOMCredito
    Set R = oDCredito.RecuperaDatosComunes(psCtaCod)
    nCalMiVivi = R!bMiVivienda
    nMontoDesemb = R!nMontoCol
    R.Close
    Set oDCredito = Nothing
    
    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oDCredDoc.RecuperaDatosBoletaDesembolso(pnMovNro, psCtaCod, psNomAge, pnRefina) 'LUCV20160621, Agregó: pnRefina Según ERS004-2016
    sCadTmp = ""
    vsCadImpresion = ""
    nNumTran = 0
    Do While Not R.EOF
        If R.Bookmark = 1 Then
            nCapital = IIf(IsNull(R!nCapital), 0, R!nCapital)
        Else
            nGasto = IIf(IsNull(R!nGasto), 0, R!nGasto)
        End If
        nNumTran = nNumTran + R!nNumTran
        R.MoveNext
    Loop
        
    R.MoveFirst
'    Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCmac)
    'ARCV 30-07-2006
    'JUEZ 20131021 *******************************************************************************
    If nEstadoCampDatos = 1 Then
        vsCadImpresion = vsCadImpresion & lsSaltoLin & lsNegritaOn
        vsCadImpresion = vsCadImpresion & "                            * *            " & Space(9) & "                            * *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "                          *     *          " & Space(9) & "                          *     *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "                         *        *        " & Space(9) & "                         *        *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "      * *               *          *       " & Space(9) & "      * *               *          *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "    *     *      *   * *            *      " & Space(9) & "    *     *      *   * *            *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "   *        *  *   *                  *    " & Space(9) & "   *        *  *   *                  *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "  *           *                        *   " & Space(9) & "  *           *                        *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & " *                                       * " & Space(9) & " *                                       *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "*                                         *" & Space(9) & "*                                         *" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO!!!" & Space(9) & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO!!!" & lsNegritaOff & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
    End If
    'END JUEZ ************************************************************************************
    lsTItular = IIf(IsNull(R!cPersNombre), "", R!cPersNombre) 'ALPA 20121226
    Call ImprimeCabeceraBoletaPago_NEW(psNomAge, R!sMoneda, "", _
            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCMAC, R!cDNI, pnAmpliado)

    'CTI3 LLENAR NUEVA CABECERA-------
    'MatAmpliadoBol
    sCadDatosTmp(1) = psNomAge 'aGENCIA
    sCadDatosTmp(2) = R!sMoneda 'TIPO DE MONEDA
    sCadDatosTmp(3) = R!FechaTran 'Fecha Tramite
    sCadDatosTmp(4) = R!HoraTran 'Hora Tramite
    sCadDatosTmp(5) = R!cPersNombre ' Nombre del Cliente
    sCadDatosTmp(6) = nNumTran 'Numero de tramite
    sCadDatosTmp(7) = psCtaCod 'Cuenta actual
    sCadDatosTmp(8) = psNomCmacAbrev 'Nombre Abreviado de la caja
    sCadDatosTmp(9) = psCodCMAC 'Cod CMC
    sCadDatosTmp(10) = R!cDNI 'DNI del Cliente
    sCadDatosTmp(17) = R!sUsuario
    'CTI3-----------------------------

    'Cuerpo de la Boleta
    
    sCadTmp = lsSaltoLin & lsSaltoLin & vsCadImpresion & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOn
    
If psCodCMAC = "102" Then  ' LIMA FEBM
    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
    Else
            sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-") & lsSaltoLin
    End If
    
    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital, 12) & lsSaltoLin
    'ARCV 30-01-2007
    If Abs(nGasto) > 0 Then
        sCadTmp = sCadTmp & "Gastos:" & oImpre.ImpreFormat(Abs(nGasto), 29) & lsSaltoLin
    End If
    '-------------
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
    sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    dProxFecPag = R!dVenc
    R.Close
    
Else   ' Liquidación Trujillo
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
    lsDescripcionBoleta1 = "DESEM LEASING"
    lsDescripcionBoleta2 = "DESEM LEASING"
    Else
    lsDescripcionBoleta1 = "DESEMB CON ABONO A CUENTA"
    lsDescripcionBoleta2 = "DESEMBOLSO  EN  EFECTIVO"
    End If
    If (pnMontoAbono > 0 Or pnAhoMontoApe > 0) Or Mid(psCtaCod, 6, 3) = "517" Then 'Modificado por BRGO 20111205
           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-")
           sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
           'CTI3 - AGREGAR
           'If (pnMontoAbono > 0 And pnAhoMontoApe = 0) Then 'Comento JOEP20190412
           If (pnMontoAbono > 0 And pnAhoMontoApe = 0) Then
                If psctaAho <> "" Then 'JOEPJOEP20190412
                    sCadDatosTmp(0) = ""
                Else
                    sCadDatosTmp(0) = 3
                End If
           End If
           '------------------------------
    Else
        sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(9, "-")
        sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(9, "-") & lsSaltoLin
             'CTI3
             If pnAmpliado = 4 Then
                sCadDatosTmp(0) = ""
             End If
    End If

    
    'sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oimpre.ImpreFormat(nCapital + pnTIF, 12)
     If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        mCalc = nCapital + pnTIF
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Monto del Desembolso" & Space(sEspa + 6 - Len("Monto del Desembolso") - Len(mCalc)) & mCalc, 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Monto del Desembolso" & Space(sEspa + 6 - Len("Monto del Desembolso") - Len(mCalc)) & mCalc, 51, 0) & lsSaltoLin
        'sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital + pnTIF, 12)
        'sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital + pnTIF, 12) & lsSaltoLin
    Else
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Monto del Desembolso" & Space(sEspa + 10 - Len("Monto del Desembolso") - Len(Format(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), "#0.00"))) & Format(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), "#0.00"), 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Monto del Desembolso" & Space(sEspa + 10 - Len("Monto del Desembolso") - Len(Format(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), "#0.00"))) & Format(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), "#0.00"), 51, 0) & Chr(13) ' & lsSaltoLin & lsSaltoLin & Chr(27)
'        sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), 12)
'        sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital - IIf(Mid(psCtaCod, 6, 3) = "517", nGasto, 0), 12) & lsSaltoLin
    End If
    
    'ARCV 30-01-2007
    If Abs(nGasto) > 0 Then
        sCadTmp = sCadTmp & "Gastos:" & oImpre.ImpreFormat(Abs(nGasto), 29)
        sCadTmp = sCadTmp & Space(14) & "Gastos :" & oImpre.ImpreFormat(Abs(nGasto), 28) & lsSaltoLin
    End If
    '-------
    
    'ARCV 24-01-2007 + 30-01-2007
    If Abs(pnMontoPoliza) > 0 Then
    '    sCadTmp = sCadTmp & "Poliza:" & oImpre.ImpreFormat(Abs(pnMontoPoliza), 29)
    '    sCadTmp = sCadTmp & Space(14) & "Poliza :" & oImpre.ImpreFormat(Abs(pnMontoPoliza), 28) & lsSaltoLin
        sCadTmp = sCadTmp & "Poliza:" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnMontoPoliza), "#0.00") & "(-)", 15)
        sCadTmp = sCadTmp & Space(2) & "Poliza :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnMontoPoliza), "#0.00") & "(-)", 28) & lsSaltoLin
    End If
    '---------------
    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
        sCadTmp = sCadTmp & lsSaltoLin
    Else
        'sCadTmp = sCadTmp & "ITF   :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnTIF), "#0.00") & "(-)", 14)
        'sCadTmp = sCadTmp & Space(2) & "ITF    :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnTIF), "#0.00") & "(-)", 28) & lsSaltoLin
        
        sCadTmp = sCadTmp & oImpre.ImpreFormat("ITF" & Space(sEspa + 10 - Len("ITF") - Len(Format(Abs(pnTIF), "#0.00") & "(-)")) & Format(Abs(pnTIF), "#0.00") & "(-)", 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("ITF" & Space(sEspa + 10 - Len("ITF") - Len(Format(Abs(pnTIF), "#0.00") & "(-)")) & Format(Abs(pnTIF), "#0.00") & "(-)", 51, 0) & Chr(13) ' & lsSaltoLin & lsSaltoLin & Chr(27)
        
    End If
    
    'ITF
    'sCadTmp = sCadTmp & lsSaltoLin
    
    
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(45, "-")
    sCadTmp = sCadTmp & Space(6) & String(44, "-") & lsSaltoLin
    'sCadTmp = sCadTmp & "Total a Desembolsar :" & Space(8) & oimpre.ImpreFormat(nCapital + NGasto, 12)
    'sCadTmp = sCadTmp & Space(15) & "Total a Desembolsar  :" & Space(8) & oimpre.ImpreFormat(nCapital + NGasto, 12) & lsSaltoLin
    
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
            sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 10)
            sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
        Else
            sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 10)
            sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
        End If
    Else
        If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
            sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 10)
            sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
        Else
 'CTI3
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total" & Space(sEspa + 10 - Len("Total") - Len(Format(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, "##,##0.00"))) & Format(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, "##,##0.00"), 51, 0) 'cti3
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total" & Space(sEspa + 10 - Len("Total") - Len(Format(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, "##,##0.00"))) & Format(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, "##,##0.00"), 51, 0) & Chr(13) ' & lsSaltoLin 'cti3
   
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Usuario" & Space(sEspa + 10 - Len("Usuario") - Len(psCodUsu)) & psCodUsu, 51, 0)
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Usuario" & Space(sEspa + 10 - Len("Usuario") - Len(psCodUsu)) & psCodUsu, 51, 0) & Chr(13) '& lsSaltoLin

            'sCadTmp = sCadTmp & "Total :" & Space(8) & oImpre.ImpreFormat(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, 10)
            'sCadTmp = sCadTmp & Space(15) & "Total :" & Space(8) & oImpre.ImpreFormat(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
            sCadTmp = sCadTmp & lsSaltoLin & lsSaltoLin
        
'            sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, 10)
'            sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + IIf(Mid(psCtaCod, 6, 3) = "517", 0, nGasto) - pnTIF - pnMontoPoliza, 12) & lsSaltoLin
'            sCadTmp = sCadTmp & lsNegritaOff
        End If
    End If
    'COMENTADO POR CTI3
    ' sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 12)
    'sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oimpre.ImpreFormat(nCapital + pnTIF, 12)
    'sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oimpre.ImpreFormat(nCapital + pnTIF, 13) & lsSaltoLin
    'sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 13) & lsSaltoLin
    'sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
    'sCadTmp = sCadTmp & lsSaltoLin
    'AMDO 20130315 CrediPermiazo ********************************
    Dim oVarSist As COMDConstSistema.DCOMGeneral
    Set oVarSist = New COMDConstSistema.DCOMGeneral
    Dim pCrediPremiazo As Integer
    pCrediPremiazo = oVarSist.LeeConstSistema(432)
    Set oVarSist = Nothing
       
    If pCrediPremiazo = 1 Then
        Dim oCredito As COMNCredito.NCOMCredito
        Set oCredito = New COMNCredito.NCOMCredito
        Dim puntosGanados As Integer
        puntosGanados = oCredito.DevuelveCantPuntosDesembolso(nCapital, psCtaCod, pdFecSis)
        Dim puntoTotales As Integer
        puntoTotales = oCredito.DevuelveCantPuntosTotalCredito(psCtaCod)
        Set oCredito = Nothing
        If puntoTotales > 0 Then
            '' AMDO20130607 Se adecuo el voucher del CrediPremiazo segun el TI-ERS075-2013
            sCadTmp = sCadTmp & "******** F E L I C I T A C I O N E S  ********" & Space(7) & "******** F E L I C I T A C I O N E S  ********"
            sCadTmp = sCadTmp & lsSaltoLin & "**" & Space(42) & "**" & Space(7) & "**" & Space(42) & "**"
            sCadTmp = sCadTmp & lsSaltoLin & "** Usted ha ganado " & puntosGanados & " opciones, Siga pagando **" & Space(7) & "** Usted ha ganado " & puntosGanados & " opciones, Siga pagando **"
            sCadTmp = sCadTmp & lsSaltoLin & "** puntualmente sus cuotas y ganarás más    **" & Space(7) & "** puntualmente sus cuotas y ganarás más    **"
            sCadTmp = sCadTmp & lsSaltoLin & "** opciones para participar en los sorteos  **" & Space(7) & "** opciones para participar en los sorteos  **"
            sCadTmp = sCadTmp & lsSaltoLin & "** de CrediPremiazo." & Space(24) & "**" & Space(7) & "** de CrediPremiazo." & Space(24) & "**"
            sCadTmp = sCadTmp & lsSaltoLin & "**" & Space(42) & "**" & Space(7) & "**" & Space(42) & "**"
            sCadTmp = sCadTmp & lsSaltoLin & "** Opciones acumuladas a la fecha: " & oImpre.Justifica(CStr(puntoTotales), 9) & "**" & Space(7) & "** Opciones acumuladas a la fecha: " & oImpre.Justifica(CStr(puntoTotales), 9) & "**"
            sCadTmp = sCadTmp & lsSaltoLin & "**********************************************" & Space(7) & "**********************************************"
            sCadTmp = sCadTmp & lsSaltoLin & lsSaltoLin 'JUEZ 20131021
        End If
    End If
    'END AMDO****************************************************
    dProxFecPag = R!dVenc
    R.Close

End If
 '' Fin
'CTI3
If sCadDatosTmp(0) = "" Or pnAmpliado = 4 Then
    MatTempDesem(0, 0) = "1"
    MatTempDesem(0, 1) = sCadTmp
End If
nTotReg = nTotReg + 1
    
    'MatTempDesem(0, 0) = "1"
    'MatTempDesem(0, 1) = sCadTmp
    'nTotReg = nTotReg + 1
    
    '*****************************************************************************
    '********** Recuperar Boletas de Gastos de Desembolso
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaGastosDesembolso(pnMovNro, psCtaCod, psNomAge)
    If Not R.BOF And Not R.EOF Then
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC)
        'Cuerpo de la Boleta
        
        sCadTmp = lsSaltoLin & lsSaltoLin & vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        If psCodCMAC = "102" Then
            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-") & lsSaltoLin
        Else
            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-")
            sCadTmp = sCadTmp & Space(7) & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(4, "-") & lsSaltoLin
        End If
    End If
    AcumTmp = 0
    Do While Not R.EOF
     If psCodCMAC = 102 Then
        sCadTmp = sCadTmp & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
     Else
        sCadTmp = sCadTmp & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12)
        sCadTmp = sCadTmp & Space(11) & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
     End If
        AcumTmp = AcumTmp + Abs(R!nMonto)
        R.MoveNext
    Loop
    
    If R.RecordCount > 0 Then
        nLin = 7 - R.RecordCount
        If nLin < 0 Then
            nLin = 0
        End If
        For i = 1 To nLin
            sCadTmp = sCadTmp & lsSaltoLin
        Next i
        If psCodCMAC = "102" Then
            sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
            R.MovePrevious
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
            sCadTmp = sCadTmp & lsSaltoLin
       Else
            sCadTmp = sCadTmp & String(40, "-")
            sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12)
            sCadTmp = sCadTmp & Space(11) & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
            R.MovePrevious
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
            sCadTmp = sCadTmp & lsSaltoLin
        End If
    End If
    R.Close
    'CTI3
    If sCadDatosTmp(0) = "" Then
        MatTempDesemGastos(0, 0) = "1"
        MatTempDesemGastos(0, 1) = sCadTmp
    End If
    nTotReg = nTotReg + 1
        
    'MatTempDesemGastos(0, 0) = "1"
    'MatTempDesemGastos(0, 1) = sCadTmp
    'nTotReg = nTotReg + 1
    '*****************************************************************************
    '********** Recuperar Boletas de Cancelacion de Creditos
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaCancelCredDesembolso(pnMovNro, psCtaCod, psNomAge)
    ReDim MatTempDesemCancel(R.RecordCount, 2)
    Do While Not R.EOF
    sCadTmp = ""
    vsCadImpresion = ""
    lsTItular = R!cPersNombre 'ALPA 20121226
    If psCodCMAC = "102" Then '' Liquidacion lima
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsNegritaOff
        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12) & lsSaltoLin
        'sCadTmp = sCadTmp & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        nTotReg = nTotReg + 1
        R.MoveNext
    Else
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC, , , nCancelado)
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-")
        sCadTmp = sCadTmp & Space(13) & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Capital" & Space(sEspa + 8 - Len("Capital") - Len(Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "##,##0.00"))) & Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "##,##0.00"), 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Capital" & Space(sEspa + 8 - Len("Capital") - Len(Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "##,##0.00"))) & Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "##,##0.00"), 51, 0) & Chr(13) '& lsSaltoLin

        sCadTmp = sCadTmp & oImpre.ImpreFormat("Interes" & Space(sEspa + 8 - Len("Interes") - Len(Format(IIf(IsNull(R!nInteres), 0, R!nInteres), "##,##0.00"))) & Format(IIf(IsNull(R!nInteres), 0, R!nInteres), "##,##0.00"), 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Interes" & Space(sEspa + 8 - Len("Interes") - Len(Format(IIf(IsNull(R!nInteres), 0, R!nInteres), "##,##0.00"))) & Format(IIf(IsNull(R!nInteres), 0, R!nInteres), "##,##0.00"), 51, 0) & Chr(13)  '& lsSaltoLin

        sCadTmp = sCadTmp & oImpre.ImpreFormat("Comisiones" & Space(sEspa + 8 - Len("Comisiones") - Len(Format(IIf(IsNull(R!comision), 0, R!comision), "##,##0.00"))) & Format(IIf(IsNull(R!comision), 0, R!comision), "##,##0.00"), 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Comisiones" & Space(sEspa + 8 - Len("Comisiones") - Len(Format(IIf(IsNull(R!comision), 0, R!comision), "##,##0.00"))) & Format(IIf(IsNull(R!comision), 0, R!comision), "##,##0.00"), 51, 0) & Chr(13) '& lsSaltoLin
        
'        sCadTmp = sCadTmp & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12)
'        sCadTmp = sCadTmp & Space(16) & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
        
'        sCadTmp = sCadTmp & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12)
'        sCadTmp = sCadTmp & Space(16) & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
                
        'CTI3 - dATOS DE CANCELACION
        sCadDatosTmp(11) = R!nCapital 'MONTO DE CAPITAL CANCELADO
        sCadDatosTmp(12) = R!cCtaCod  'Cuenta anterior Cancelada
        'sCadTmp = sCadTmp & Space(16) & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
                
        'ARCV 30-01-2007
        If R!nGastos > 0 Then
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Gastos" & Space(sEspa + 8 - Len("Gastos") - Len(Format(IIf(IsNull(R!nGastos), 0, R!nGastos), "##,##0.00"))) & Format(IIf(IsNull(R!nGastos), 0, R!nGastos), "##,##0.00"), 51, 0) 'cti3
            sCadTmp = sCadTmp & oImpre.ImpreFormat("Gastos" & Space(sEspa + 8 - Len("Gastos") - Len(Format(IIf(IsNull(R!nGastos), 0, R!nGastos), "##,##0.00"))) & Format(IIf(IsNull(R!nGastos), 0, R!nGastos), "##,##0.00"), 51, 0) & Chr(13) '& lsSaltoLin
            'sCadTmp = sCadTmp & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12)
            'sCadTmp = sCadTmp & Space(16) & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
        End If
        '-------------
        If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
            sCadTmp = sCadTmp & lsSaltoLin
        Else
            sCadTmp = sCadTmp & oImpre.ImpreFormat("ITF" & Space(sEspa + 8 - Len("ITF") - Len(Format(IIf(IsNull(R!nITF), 0, R!nITF), "##,##0.00"))) & Format(IIf(IsNull(R!nITF), 0, R!nITF), "##,##0.00"), 51, 0) 'cti3
            sCadTmp = sCadTmp & oImpre.ImpreFormat("ITF" & Space(sEspa + 8 - Len("ITF") - Len(Format(IIf(IsNull(R!nITF), 0, R!nITF), "##,##0.00"))) & Format(IIf(IsNull(R!nITF), 0, R!nITF), "##,##0.00"), 50, 0) & Chr(13) '& lsSaltoLin
            'sCadTmp = sCadTmp & "ITF     :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nITF), 0, R!nITF), 12)
            'sCadTmp = sCadTmp & Space(16) & "ITF     :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nITF), 0, R!nITF), 12) & lsSaltoLin
        End If
        
        Dim nTotal As Double
        nTotal = (R!nCapital + IIf(IsNull(R!nInteres) = True, 0, R!nInteres) + IIf(IsNull(R!comision) = True, 0, R!comision) + IIf(IsNull(R!nGastos) = True, 0, R!nGastos)) + IIf(IsNull(R!nITF) = True, 0, R!nITF)
        'sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & String(40, "-")
        sCadTmp = sCadTmp & Space(11) & String(40, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Total a Pagar" & Space(sEspa + 8 - Len("Total a Pagar") - Len(Format(IIf(IsNull(nTotal), 0, nTotal), "##,##0.00"))) & Format(IIf(IsNull(nTotal), 0, nTotal), "##,##0.00"), 51, 0) 'cti3
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Total a Pagar" & Space(sEspa + 8 - Len("Total a Pagar") - Len(Format(IIf(IsNull(nTotal), 0, nTotal), "##,##0.00"))) & Format(IIf(IsNull(nTotal), 0, nTotal), "##,##0.00"), 51, 0) & Chr(13)  'cti3 & lsSaltoLin
        
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Usuario" & Space(sEspa + 8 - Len("Usuario") - Len(psCodUsu)) & psCodUsu, 51, 0)
        sCadTmp = sCadTmp & oImpre.ImpreFormat("Usuario" & Space(sEspa + 8 - Len("Usuario") - Len(psCodUsu)) & psCodUsu, 51, 0) & Chr(13) ' & lsSaltoLin
        
'        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nITF), 0, R!nITF), 12)
'        sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nITF), 0, R!nITF), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsNegritaOff
'        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12)
'        sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12) & lsSaltoLin
        'sCadTmp = sCadTmp & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        sCadTmp = sCadTmp & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        
        If sCadDatosTmp(0) = "" Then
            MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
            MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        End If
        'MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
        'MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        nTotReg = nTotReg + 1
        R.MoveNext
    End If
        
    Loop
    R.Close
    
    '***********************************************************
    'Recuperar Boletas de Apertura de Cuenta
    '***********************************************************
    Dim oFecha As COMDConstSistema.DCOMGeneral
    Set oFecha = New COMDConstSistema.DCOMGeneral
    If pnAhoMontoApe > 0 Then
        MatTempAhoApe(0, 0) = "1"
        MatTempAhoApe(0, 1) = EmiteBoletaAhorro(psctaAho, pnAhoMontoApe, pnAhoMontoApe, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), psNomAge, psCodUsu, psCodCMAC, , pnTIF)
        nTotReg = nTotReg + 1
        'CTI3 - AGREGAR
         sCadDatosTmp(0) = 3
        '------------------------------
    End If
    
    
    '***********************************************************
    'Recuperar Boletas de Abono de Cuenta
    '***********************************************************
    Dim sCadBoletaAbono As String
    Dim C As COMDConstSistema.FCOMITF
    Dim nItfValor As Double
    Dim nRedondeoITF As Double
    Dim nITFMonto As Double
    
    Set C = New COMDConstSistema.FCOMITF
    Set oNCred = New NCOMCredito 'BRGO 20111022
    Call C.fgITFParametros
    nItfValor = C.gnITFPorcent
    
    'CTI3
    sCadDatosTmp(13) = psctaAho  'Cuenta de Ahorro donde se Realiza el Abono
    sCadDatosTmp(14) = pnMontoAbono 'Monto Abonado
    sCadDatosTmp(15) = pnTIF 'Monto ITF
    '----------------
    
    If Mid(psCtaCod, 6, 3) <> "517" Then 'BRGO 20111125
        If pnMontoAbono > 0 And pnAhoMontoApe <= 0 Then
            MatTempAhoApe(0, 0) = "1"
            MatTempAhoApe(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoAbono, pnSaldoDispAhoAbono, _
                pnSaldoCntAhoAbono, pnIntGanadoAbono, 1, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
                psNomAge, psCodUsu, psCodCMAC, pnTIF)
            nTotReg = nTotReg + 1
        End If 'END BRGO
    Else
        If pnRefina <> 1 Then 'LUCV20161110, Agregó: pnRefina <> 1 - En creditos refinanciado no aplica
            MatTempAhoApe(0, 0) = "1"
            For i = 0 To 5
                'nITFMonto = C.fgITFCalculaImpuesto(pMatDatosAhoEmp(i * 14 + 8)) 'BRGO 20111125
                nITFMonto = C.fgITFCalculaImpuesto(pMatDatosAhoEmp(i * 14 + 1)) 'EJVG20130322
                nRedondeoITF = oNCred.fgDiferenciaRedondeoITF(nITFMonto) 'BRGO 20111125
                nITFMonto = IIf(nRedondeoITF > 0, nITFMonto - nRedondeoITF, nITFMonto) 'BRGO 20111125
                    
                sCadBoletaAbono = sCadBoletaAbono & EmiteBoletaAhorroAbonoRetiro(pMatDatosAhoEmp(i * 14 + 0), pMatDatosAhoEmp(i * 14 + 1), pMatDatosAhoEmp(i * 14 + IIf(i <> 4, 8, 10)), _
                    pMatDatosAhoEmp(i * 14 + IIf(i <> 4, 9, 11)), pMatDatosAhoEmp(i * 14 + 2), IIf(i <> 4, 1, 2), Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
                    psNomAge, psCodUsu, psCodCMAC, nITFMonto, False, psCtaCod, lsTItular)
                    'ALPA 20121226 Se agregó lsTItular
                If i = 0 Or i = 3 Then
                    sCadBoletaAbono = sCadBoletaAbono & lnSaltoLinDoc
                    sCadBoletaAbono = sCadBoletaAbono & Chr$(12)
                End If
            Next
            MatTempAhoApe(0, 1) = sCadBoletaAbono & psBolTransfGarInfoGas
            'nTotReg = nTotReg + 1 'LUCV20161110, Comentó
        End If
             nTotReg = nTotReg + 1 'LUCV20161110
    End If
    
        '********CTI3
        If sCadDatosTmp(0) = "3" Then
            MatTempAhoApe(0, 1) = ""
        End If
    
    '***********************************************************
    'Recuperar Boletas de Retiro Gastos de Cuenta
    '***********************************************************
    If pnMontoRetGastos > 0 Then
        MatTempAhoRetGastos(0, 0) = "1"
        nITFMonto = C.fgITFCalculaImpuesto(pnMontoRetGastos) 'BRGO 20111022
        nRedondeoITF = oNCred.fgDiferenciaRedondeoITF(nITFMonto) 'BRGO 20111022
        nITFMonto = IIf(nRedondeoITF > 0, nITFMonto - nRedondeoITF, nITFMonto) 'BRGO 20111022
                
        MatTempAhoRetGastos(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetGastos, pnSaldoDispAhoRetGastos, _
            pnSaldoCntAhoRetGastos, pnIntGanadoRetGastos, 2, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC, nITFMonto) 'C.fgITFCalculaImpuesto(pnMontoRetGastos))'BRGO 20111022
        nTotReg = nTotReg + 1
    End If
    
    '***********************************************************
    'Recuperar Boletas de Retiro Cancelaciones de Cuenta
    '***********************************************************
    'CTI3 - Monto de Cancelacion
     sCadDatosTmp(16) = pnSaldoDispAhoRetCancel 'Monto Total a Entregar
    '--------------------------
    If pnMontoRetCancel > 0 Then
        MatTempAhoRetCancel(0, 0) = "1"
        'JUEZ 20130930 Comentado para no imprimir ITF ******************************************
        'nITFMonto = C.fgITFCalculaImpuesto(pnMontoRetCancel) 'BRGO 20111022
        'nRedondeoITF = oNCred.fgDiferenciaRedondeoITF(nITFMonto) 'BRGO 20111022
        'nITFMonto = IIf(nRedondeoITF > 0, nITFMonto - nRedondeoITF, nITFMonto) 'BRGO 20111022
        nITFMonto = 0
        'END JUEZ ******************************************************************************
        
        MatTempAhoRetCancel(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetCancel, pnSaldoDispAhoRetCancel, _
            pnSaldoCntAhoRetCancel, pnIntGanadoRetCancel, 2, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC, nITFMonto) 'C.fgITFCalculaImpuesto(pnMontoRetCancel))'BRGO 20111022
        nTotReg = nTotReg + 1
    End If
    
        '********CTI3
        If sCadDatosTmp(0) = "3" Then
            MatTempAhoRetCancel(0, 1) = ""
        End If
    
    '*** FRHU 20140228 RQ14006
    '***********************************************************
    'Recuperar Boletas de Retiro Efectivo de Cuenta
    '***********************************************************
    If pnMontoRetiroEfectivo > 0 Then
        MatTempAhoRetEfectivo(0, 0) = "1"
        'JUEZ 20130930 Comentado para no imprimir ITF ******************************************
        nITFMonto = C.fgITFCalculaImpuesto(pnMontoRetiroEfectivo) 'BRGO 20111022
        nRedondeoITF = oNCred.fgDiferenciaRedondeoITF(nITFMonto) 'BRGO 20111022
        nITFMonto = IIf(nRedondeoITF > 0, nITFMonto - nRedondeoITF, nITFMonto) 'BRGO 20111022
        'nITFMonto = 0
        'END JUEZ ******************************************************************************
        
        MatTempAhoRetEfectivo(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetiroEfectivo, pnSaldoDispAhoRetEfectivo, _
            pnSaldoCntAhoRetEfectivo, pnIntGanadoRetEfectivo, 2, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC, nITFMonto) 'C.fgITFCalculaImpuesto(pnMontoRetCancel))'BRGO 20111022
        nTotReg = nTotReg + 1
    End If
    'FIN FRHU 20140228 RQ14006
    
    'Set oFecha = Nothing
    
    Set C = Nothing
    
    '********************************************************************************
    'Recuperar Documento de Comprobante de Desembolso
    '********************************************************************************
    'LUCV20160621, Según ERS004-2016 ->**********
    If pnRefina = 1 Then
    MatTempComprob(0, 0) = "2"
    MatTempComprob(0, 1) = ImprimeComprobanteVigencia(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, psNomCmac, pnRefina)
    nTotReg = nTotReg + 1
    'Fin LUCV20160621 <-**********
    Else
    MatTempComprob(0, 0) = "2"
    MatTempComprob(0, 1) = ImprimeComprobanteDesembolso(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, psNomCmac)
    nTotReg = nTotReg + 1
    End If
    '**********************CTI3 - BOLETA AMPLIADO UNIFICADO**************************
    If sCadDatosTmp(0) = "3" Then
        Call ImprimeCabeceraBoletaAmpliado(sCadDatosTmp)
        nTotReg = nTotReg + 1
    End If
    '*******************************************************
    '********************************************************************************
    'Recuperar Documento de Plan de Pagos
    '********************************************************************************
    Dim loRep As COMNCredito.NCOMCalendario
    Set loRep = New COMNCredito.NCOMCalendario
    MatTempPlan(0, 0) = "2"
    MatTempPlan(0, 1) = ""
    'For I = 1 To 3
    'ARCV 01-11-2006
    
    '*** PEAC 20080723
    'For i = 1 To 2
        MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, , psNomCmac, , , True)
        'WIOR 20160111 - COMENTO
        '        If nCalMiVivi = True Then
        '            MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, True, , , , True)
        '        End If
    'Next i
    '*** FIN PEAC
    
    nTotReg = nTotReg + 1
    
    'CTI5 20210516***
    lsBoletaApeFMV = ""
    If (Mid(psCtaCod, 6, 3) = "802" Or Mid(psCtaCod, 6, 3) = "806") Then
        If Not IsArray(pMatCuentaGastoCierre) Then
            pMatCuentaGastoCierre = oDCredDoc.ObtenerMatrizGastoCierrexReimprDesemb(pnMovNro)
        End If
    
        If IsArray(pMatCuentaGastoCierre) Then
            lsBoletaApeFMV = EmiteBoletaAhorroAbonoRetiro(pMatCuentaGastoCierre(12), pMatCuentaGastoCierre(7), pMatCuentaGastoCierre(7) - pMatCuentaGastoCierre(13), pMatCuentaGastoCierre(7) - pMatCuentaGastoCierre(13), 0, 1, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), psNomAge, psCodUsu, psCodCMAC, pMatCuentaGastoCierre(13))
            nTotReg = nTotReg + 1
        End If
    End If
    
    Set oFecha = Nothing
    'END CTI5*******
    
    ReDim MatDocs(nTotReg, 2)
    MatDocs(0, 0) = MatTempDesem(0, 0)
    MatDocs(0, 1) = MatTempDesem(0, 1)
    MatDocs(1, 0) = MatTempDesemGastos(0, 0)
    MatDocs(1, 1) = MatTempDesemGastos(0, 1)
    nTotReg = 2
    For i = 0 To UBound(MatTempDesemCancel) - 1
        MatDocs(i + 2, 0) = MatTempDesemCancel(i, 0)
        MatDocs(i + 2, 1) = MatTempDesemCancel(i, 1)
        nTotReg = nTotReg + 1
    Next i
    'ARCV 01-04-2007   : Cambiar de Posicion
    '    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
    '    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
    '    nTotReg = nTotReg + 1
    '----------
    'Apertura
    If pnAhoMontoApe > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Abono
    If (pnMontoAbono > 0 And pnAhoMontoApe <= 0) Or (Mid(psCtaCod, 6, 3) = "517") Then 'BRGO 20111125
        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Retiro de Gastos
     If pnMontoRetGastos > 0 Or Mid(psCtaCod, 6, 3) = "517" Then
        MatDocs(nTotReg, 0) = MatTempAhoRetGastos(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoRetGastos(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Retiro de Cancelaciones
    If pnMontoRetCancel > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoRetCancel(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoRetCancel(0, 1)
        nTotReg = nTotReg + 1
    End If
    
    'FRHU20140228 RQ14006
    'Retiro de Efectivo
    If pnMontoRetiroEfectivo > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoRetEfectivo(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoRetEfectivo(0, 1)
        nTotReg = nTotReg + 1
    End If
    'FIN FRHU20140228 RQ14006
    
    'CTI5 20210516***
    If Len(Trim(lsBoletaApeFMV)) > 0 Then
        MatDocs(nTotReg, 0) = "1"
        MatDocs(nTotReg, 1) = lsBoletaApeFMV
        nTotReg = nTotReg + 1
    End If
    'END CTI5 *******
    
    'ARCV 01-04-2007: Cambiar de Posicion
    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
    nTotReg = nTotReg + 1
    '------------
    
    MatDocs(nTotReg, 0) = MatTempPlan(0, 0)
    MatDocs(nTotReg, 1) = MatTempPlan(0, 1)
    
    'CTI3
    nTotReg = nTotReg + 1
    vsCadImpreAmpli = vsCadImpreAmpli + sCadTmp
    MatDocs(nTotReg, 0) = sCadDatosTmp(0)
    MatDocs(nTotReg, 1) = vsCadImpreAmpli
    'en caso sea un deposito en cuenta
    If sCadDatosTmp(0) = "3" Then
        For i = 0 To UBound(MatDocs) - 1
          If MatDocs(i, 0) = "1" And (i <> 3 And pnAhoMontoApe > 0) Then
            MatDocs(i, 0) = ""
          End If
        Next i
    End If
    
    ImprimeDocumentosDesembolso = MatDocs
    
End Function

Public Function ImprimePlandePagos(ByVal psCtaCod As String, ByVal psNomAge As String, _
    ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
    ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0, _
    Optional ByVal pbDetallado As Boolean = False, _
    Optional ByVal pbMostrarITF As Boolean = True, _
    Optional ByVal psFecDes As String, _
    Optional ByVal pbDuplicado As Boolean = False) As String 'ARCV 08-06-2006
    '************Parametro psFecDes agregado por PASI20131127 segun TI-ERS136-2013
    'JUEZ 20151229 Se agregó parámetro pbDuplicado

Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset 'ANGC/RIRO 20200520
Dim oCredDoc2 As COMDCredito.DCOMCredDoc 'ANGC/RIRO 20200520
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim oCalend As COMDCredito.DCOMCalendario
Dim oNCred As COMNCredito.NCOMCredito
Dim MatCalend As Variant
Dim sCadImp As String
Dim i As Integer
Dim nIntereses As Double
Dim nCapital As Double
Dim oParam As COMDCredito.DCOMParametro
Dim nTramoNoConsPorcen As Double
Dim RDesPar As ADODB.Recordset
Dim bEstadoConvenio As Boolean
Dim nDiasGracia As Integer 'agregado por PASI20131127 segun TI-ERS136-2013

'ARCV 12-02-2007
Dim nIntGracia As Double
Dim nGastos As Double
Dim nGastosComi As Double
Dim nGastosSegDes As Double
Dim nGastosEstCta As Double

Dim nGastosPolizaIncendio As Double '*** PEAC 20090721
Dim nMultMYPE As Double 'RECO20150601 ERS023-2015
Dim nTotCuota As Double '*** PEAC 20080722
'---------------

'ARCV 08-06-2006
Dim oITF As COMDConstSistema.FCOMITF
Dim nCuota As Double
Dim nCuotaITF As Double

Set oITF = New COMDConstSistema.FCOMITF
Call oITF.fgITFParametros

'----------------------

'ARCV 15-02-2007
Dim oCredRel As COMDCredito.UCOMCredRela
Dim bExisteMasDeUnTitular As Boolean
Dim oDCred As COMDCredito.DCOMCredActBD
Dim RSeg As ADODB.Recordset
Dim nValorGastoSeg As Double

'**DAOR 20070403************************************
Dim oParametros As COMDCredito.DCOMParametro
Dim nPorComiCorrespMN As Double
Dim nPorComiCorrespME As Double
Dim nMonMinComiCorrespMN As Double
Dim nMonMaxComiCorrespMN As Double
Dim nMonMinComiCorrespME As Double
Dim nFactorCorresp As Double
Dim nMonMinComi As Double
Dim nMonMaxComi As Double
Dim nComiCorresp As Double
Dim nTasaEfectivaAnual As Double
Dim nTasaCostoEfectivoAnual As Double
Dim MatCalendTemp As Variant
Dim lnContador As Integer '*** PEAC 20080811
Dim cPagoAnticipado As String 'JUEZ 20150505
Dim rsAmortCap As ADODB.Recordset 'JUEZ 20150505

'*************************************************
Dim rTasCosEfeAnu As ADODB.Recordset 'DAOR 20070419
'WIOR 20120528 *******************************************************
Dim bMicroseguro As Boolean
Dim bMultiriesgo As Boolean
Dim nGastosPolizaMicroseguro As Double
Dim nGastosPolizaMultiriesgo As Double
Dim oCreditoMM As COMDCredito.DCOMCredito
Dim rsCreditoMM As ADODB.Recordset
bMicroseguro = False
bMultiriesgo = False
Set oCreditoMM = New COMDCredito.DCOMCredito
Set rsCreditoMM = oCreditoMM.ObtenerMicroseguro(psCtaCod)

If rsCreditoMM.RecordCount > 0 Then
    bMicroseguro = True
End If
Set rsCreditoMM = oCreditoMM.ObtenerMultiriesgo(psCtaCod)
If rsCreditoMM.RecordCount > 0 Then
    bMultiriesgo = True
End If
Set rsCreditoMM = Nothing
Set oCreditoMM = Nothing
'WIOR FIN ************************************************************

'->*****LUCV20180601, Agregó según ERS022-2018
Dim oDCOMCalendario As COMDCredito.DCOMCalendario
Dim bVerCalendIntGrac As Boolean

Set oDCOMCalendario = New COMDCredito.DCOMCalendario
bVerCalendIntGrac = False
bVerCalendIntGrac = oDCOMCalendario.TieneCalendarioMetodoInterativo(psCtaCod)
Set oDCOMCalendario = Nothing
'<-***** Fin LUCV20180601
'WIOR 20150710 COMENTO ESTA PARTE
'Set oCredRel = New COMDCredito.UCOMCredRela
'Call oCredRel.CargaRelacPersCred(psCtaCod)
'bExisteMasDeUnTitular = oCredRel.ExisteTitularYCodeudor(psCtaCod) 'WIOR 20130419 AGREGO psCtaCod
'Set oCredRel = Nothing
'
'Set oDCred = New COMDCredito.DCOMCredActBD
'Set RSeg = oDCred.CargaRecordSet("SELECT*FROM ProductoConcepto WHERE nPrdConceptoCod=" & gColocConceptoCodGastoSeguro7)
'If RSeg.EOF Then
'    nValorGastoSeg = 0
'Else
'    nValorGastoSeg = IIf(bExisteMasDeUnTitular, RSeg!nValorDosTit, RSeg!nValor)
'End If

'WIOR 20150710 ***
Set oCredRel = New COMDCredito.UCOMCredRela
Set RSeg = oCredRel.ObtenerTitularesYTasaSegDes(psCtaCod)

If Not (RSeg.EOF And RSeg.BOF) Then
    bExisteMasDeUnTitular = IIf(CInt(RSeg!nAsegurados) > 1, True, False)
    nValorGastoSeg = RSeg!nTasa
Else
    bExisteMasDeUnTitular = False
    nValorGastoSeg = 0
End If

Set oDCred = New COMDCredito.DCOMCredActBD
'WIOR FIN ********

'**DAOR 20070419*****************************************************
Set rTasCosEfeAnu = oDCred.CargaRecordSet("select nTasCosEfeAnu from ColocacCred where cCtaCod='" & psCtaCod & "'")
If rTasCosEfeAnu.EOF Then
    nTasaCostoEfectivoAnual = 0
Else
    nTasaCostoEfectivoAnual = IIf(IsNull(rTasCosEfeAnu!nTasCosEfeAnu), 0, rTasCosEfeAnu!nTasCosEfeAnu)
End If
'********************************************************************

'JUEZ 201520505 ***************************************
Set rsAmortCap = oDCred.CargaRecordSet("exec stp_sel_VerificaAmortizacionCapital '" & psCtaCod & "'")
If rsAmortCap.RecordCount > 0 Then cPagoAnticipado = "(Remanente)"
Set rsAmortCap = Nothing
'END JUEZ *********************************************

Set oDCred = Nothing
'-------

    '**DAOR 20070331*************************************************
    nFactorCorresp = 1
    nMonMinComi = 0
    nMonMaxComi = 0
    If psCtaCod <> "" Then
        Set oParametros = New COMDCredito.DCOMParametro
        nPorComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaMN)
        nPorComiCorrespME = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaME)
        nMonMinComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaMN)
        nMonMaxComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMaxCobroCorresponsaliaMN)
        nMonMinComiCorrespME = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaME)
        If Mid$(psCtaCod, 9, 1) = "1" Then 'MN
            nFactorCorresp = nPorComiCorrespMN / 100
            nMonMinComi = nMonMinComiCorrespMN
            nMonMaxComi = nMonMaxComiCorrespMN
        Else
            nFactorCorresp = nPorComiCorrespME / 100
            nMonMinComi = nMonMinComiCorrespME
            nMonMaxComi = 9999999999#
        End If
    End If
    '***************************************************************
    
    Set oCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
    Set oCredDoc = Nothing
    
    If R.EOF Then Exit Function
    
    If pbMiViv Then
        Set oParam = New COMDCredito.DCOMParametro
        nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
        Set oParam = Nothing
        If Not pbCalendParalelo Then
            pnMontoPrestamo = pnMontoPrestamo * ((100 - nTramoNoConsPorcen) / 100)
            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
        Else
            pnMontoPrestamo = pnMontoPrestamo * (nTramoNoConsPorcen / 100)
            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
        End If
    End If
    
    Set oCalend = New COMDCredito.DCOMCalendario
    Set RDesPar = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    Set oNCred = New COMNCredito.NCOMCredito
    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod, pnMontoPrestamo, pbCalendParalelo, IIf(pbCalendParalelo, True, False), psFecDes, IIf(IsNull(R!nTasaInteres), 0, R!nTasaInteres)) 'PASI 20131219 AGREGO psFecDes, IIf(IsNull(R!nTasaInteres), 0, R!nTasaInteres))
    Set oNCred = Nothing
    
    '**DAOR 20070403**************************************
    'ReDim MatCalendTemp(UBound(MatCalend), 3)
    'For i = 0 To UBound(MatCalend) - 1
    '    MatCalendTemp(i, 0) = MatCalend(i, 0) 'Fecha de Pago
    '    MatCalendTemp(i, 1) = MatCalend(i, 1) 'Nro de cuota
    '    MatCalendTemp(i, 2) = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)) 'Monto Cuota incluyendo gastos y comisiones
    'Next
    Set oNCred = New COMNCredito.NCOMCredito
        nTasaEfectivaAnual = oNCred.TasaIntPerDias(R!nTasaInteres, 360) * 100
    '    nTasaCostoEfectivoAnual = oNCred.GeneraTasaCostoEfectivoAnual(R!nMonto, MatCalendTemp)
    Set oNCred = Nothing
    '*****************************************************
    'By Capi 28022008 para el Banco Nacion
    Dim sCuentaRefBN As String
    sCuentaRefBN = Mid(psCtaCod, 4, 3) & Mid(psCtaCod, 8, 2) & Mid(psCtaCod, 11, 7)
        
    sCadImp = ""
    '*** PEAC 20080811
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", "", lsTab, False)'LUCV20180601, Comentó según ERS022-2018
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "CRONOGRAMA DE PAGOS", "", lsTab, False) 'LUCV20180601, Agregó según ERS022-2018
    
    'sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
    If pbDuplicado Then sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10) 'JUEZ 20151229
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & "Operación AF  : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
    Else
        sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
    End If
    sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & Space(1) & "Cuota    : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
    'MAVM 20120124 ***
    'sCadImp = sCadImp & lsTab & "Tasa Interes  : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & Space(1) & "Plazo    : " & Trim(Str(R!nPlazo)) & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Mensual" & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Plazo         : " & Trim(Str(R!nPlazo)) & Chr$(10)Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "CRONOGRAMA DE PAGOS", "", lsTab, False)'LUCV20180601, Agregó según ERS022-2018
    sCadImp = sCadImp & lsTab & "Frecuencia    : " & Trim(Str(R!nPlazo)) & Chr$(10) 'LUCV20180601, Agregó según ERS022-2018
    '***
    '****************************************************
    'Modificado por PASI20131128 TIERS136-2013
    'sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
    If psFecDes = Empty Then
        sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    '->***** LUCV20180601, Agregó condición. según ERS022-2018
    If bVerCalendIntGrac Then
        sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
    Else
        sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Chr$(10)
    End If
    '<-***** Fin LUCV20180601
    Else
        sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(CDate(psFecDes), "ddd, d mmm yyyy") & Chr$(10)
        nDiasGracia = CInt(R!nPeriodoGracia) - CInt(DateDiff("D", R!dFecVig, CDate(psFecDes)))
        sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(IIf(nDiasGracia < 0, 0, nDiasGracia))) & Chr$(10)
    End If
    'END PASI********************************************
    
    If pbComodin Then
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Cuota Comodin", 30, 0)
    ElseIf pbMiViv Then
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Mi Vivienda", 30, 0)
    Else
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
    End If
    
    If pbCalendDin Then
        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("SI", 30, 0)
    Else
        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)
    End If
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual       : " & Format(nTasaEfectivaAnual, "#0.00") & " %" & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual : " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual " & cPagoAnticipado & ": " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10) 'JUEZ 20150505
    
    'RIROANGC 2020    --ObtieneDatosActivatePeru
    Set oCredDoc2 = New COMDCredito.DCOMCredDoc
    Set R2 = oCredDoc2.ObtieneDatosActivatePeru(psCtaCod)
    Set oCredDoc2 = Nothing
    Dim campanaReactivate As Boolean
    If R2.EOF And R2.BOF Then 'ANGC20200529
        campanaReactivate = False
    Else
        campanaReactivate = True
    End If
    If campanaReactivate = True Then
        sCadImp = sCadImp & lsTab & "Comision Cofide   : " & "0.5% " & "Anual" & Chr$(10)
    End If
    'END
    
    'By Capi 28022008
    If Not (RDesPar.BOF Or RDesPar.EOF) Then
        If RTrim(RDesPar!cConvenioBN) <> "NO" Then
            sCadImp = sCadImp & lsTab & "CODIGO CREDITO MYPE PAGO BANCO DE LA NACION : " & oImpresora.gPrnNegritaON & sCuentaRefBN & oImpresora.gPrnNegritaOFF & Chr$(10)
        End If
    End If
    
    If RDesPar.RecordCount >= 2 Then
        sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("DESEMBOLSO", 10, 0) & oImpre.ImpreFormat("ESTADO", 10, 2) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        Do While Not RDesPar.EOF
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & oImpre.ImpreFormat(RDesPar!nCapital, 10, 2, True) & oImpre.ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
            RDesPar.MoveNext
        Loop
    End If
    'By Capi 28022008
    'RDesPar.Close
    
    sCadImp = sCadImp & Chr$(10)
    
    'CUSCO
    If pbDetallado Then
        ''ARCV 14-02-2007
        ''sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gastos", 9, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & Chr$(10)
        ''sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 106, 90), "-") & Chr$(10)
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gast/Comi.(1)", 15, 0) & oImpre.ImpreFormat("Seg.Desgr.(2)", 15, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & Chr$(10)
        'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 122, 90), "-") & Chr$(10)
        ''------
        '**DAOR 20070403*********************************************
        
        '*** PEAC 20080722
        'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 110), "-") & Chr$(10)
        'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
        'WIOR 20120528***
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)'RECO20150601 ERS023-2015
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 104)), "-") & Chr$(10) 'RECO20150601 ERS023-2015
        sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10) 'JUEZ 20150603
        'WIOR FIN****
'       sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 17, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 14, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gast. y", 11, 0) & oImpre.ImpreFormat("Gastos", 11, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & oImpre.ImpreFormat("Pago Bco.", 12, 0) & Chr$(10)
'       sCadImp = sCadImp & lsTab & Space(17) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(47) & oImpre.ImpreFormat("Comi.(1)", 11, 0) & oImpre.ImpreFormat("Est. Cta.", 11, 0) & oImpre.ImpreFormat("Desgr.(2)", 11, 0) & Space(24) & oImpre.ImpreFormat("Naci?n(3)", 12, 0) & Chr$(10)

        '*** PEAC 20090720
'        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 11, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 14, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 8, 0) & oImpre.ImpreFormat("Gast. y", 9, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & oImpre.ImpreFormat("Pago Bco.", 12, 0) & Chr$(10)
'        sCadImp = sCadImp & lsTab & Space(11) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(33) & oImpre.ImpreFormat("Comi.(1)", 9, 0) & oImpre.ImpreFormat("Desgr.(2)", 11, 0) & Space(12) & oImpre.ImpreFormat("Naci?n(3)", 12, 0) & Chr$(10)
'        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("de Pago", 8) & Space(3) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(25) & oImpre.ImpreFormat("(1)", 6) & oImpre.ImpreFormat("Comi.(2)", 9, 0) & oImpre.ImpreFormat("Desgr.(3)", 11, 0) & oImpre.ImpreFormat("incendio", 11, 0) & Chr$(10)
        'ALPA 20120424
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast.Com.", 10, 0) & oImpre.ImpreFormat("Seg.Desg", 9, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
        'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("e IGV", 10, 0) & oImpre.ImpreFormat("Mult", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
        'WIOR 20120528***********************************
            If bMicroseguro = True Then
                If bMultiriesgo = True Then
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast.Com.", 10, 0) & oImpre.ImpreFormat("Seg.Desg", 9, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Mic-", 6, 0) & oImpre.ImpreFormat("Sg Mul", 8, 0) & oImpre.ImpreFormat("Saldo", 7, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("e IGV", 10, 0) & oImpre.ImpreFormat("Mult", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat("Seg.", 6, 0) & oImpre.ImpreFormat("Riesg.", 6, 0) & Space(3) & oImpre.ImpreFormat("Cap.", 8, 0) & Chr$(10)
                Else
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast.Com.", 10, 0) & oImpre.ImpreFormat("Seg.Desg", 9, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat(" Micro-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("e IGV", 10, 0) & oImpre.ImpreFormat("Mult", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Seg.", 10, 0) & Chr$(10)
                End If
            Else
                If bMultiriesgo = True Then
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast.Com.", 10, 0) & oImpre.ImpreFormat("Seg.Desg", 9, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Seg Mult-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("e IGV", 10, 0) & oImpre.ImpreFormat("Mult", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Riesg.", 10, 0) & Chr$(10)
                Else
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast.Com.", 10, 0) & oImpre.ImpreFormat("Seg.Desg", 9, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("e IGV", 10, 0) & oImpre.ImpreFormat("Mult", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
                End If
            End If
        'WIOR FIN ***************************************
        Else
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
        'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
        'WIOR 20120528***********************************
            If bMicroseguro = True Then
                If bMultiriesgo = True Then
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Mic-", 6, 0) & oImpre.ImpreFormat("Sg Mul", 8, 0) & oImpre.ImpreFormat("Saldo", 7, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat("Seg.", 6, 0) & oImpre.ImpreFormat("Riesg.", 6, 0) & Space(3) & oImpre.ImpreFormat("Cap.", 8, 0) & Chr$(10)
                Else
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat(" Micro-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Seg.", 10, 0) & Chr$(10)
                End If
            Else
                If bMultiriesgo = True Then
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Seg Mult-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Riesg.", 10, 0) & Chr$(10)
                Else
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & Space(3) & oImpre.ImpreFormat("Mult.", 8, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    
                    '->***** LUCV20180601, Agregó condición. según ERS022-2018
                    If bVerCalendIntGrac Then
                        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 8, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.", "Seg."), 7, 0) & oImpre.ImpreFormat("Mult.", 7, 0) & oImpre.ImpreFormat("Saldo Cap", 11, 0) & Chr$(10) 'JUEZ 20150603
                        sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Mult.", "Vehic."), 7, 0) & oImpre.ImpreFormat("PYME", 7, 0) & Chr$(10) 'JUEZ 20150724
                    Else
                        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Gast. y", 8, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.", "Seg."), 7, 0) & oImpre.ImpreFormat("Mult.", 7, 0) & oImpre.ImpreFormat("Saldo Cap", 11, 0) & Chr$(10)
                        sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Mult.", "Vehic."), 7, 0) & oImpre.ImpreFormat("PYME", 7, 0) & Chr$(10) 'JUEZ 20150724
                    End If
                    'Fin LUCV0180601
                    'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
                    'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & oImpre.ImpreFormat("MYPE.(2)", 10, 0) & Chr$(10)
                    'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Com.(1)", 8, 0) & oImpre.ImpreFormat("Desg(2)", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("MYPE", 7, 0) & Chr$(10) 'JUEZ 20150603
                    'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("MYPE", 7, 0) & Chr$(10) 'JUEZ 20150724
                End If
            End If
        'WIOR FIN ***************************************
        End If
        
        'sCadImp = sCadImp & lsTab & Space(11) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(25) & oImpre.ImpreFormat("(1)", 6) & oImpre.ImpreFormat("Comi.(2)", 9, 0) & oImpre.ImpreFormat("Desgr.(3)", 11, 0) & oImpre.ImpreFormat("incendio", 11, 0) & Chr$(10)
        
        'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 110), "-") & Chr$(10)
        'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
        '*** FIN PEAC
        'WIOR 20120528***
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)'RECO20150601 ERS023-2015
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 104)), "-") & Chr$(10) 'RECO20150601 ERS023-2015
        sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10) 'JUEZ 20150603
        'WIOR FIN****
        '************************************************************
    Else
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 25, 0) & oImpre.ImpreFormat("No Cuota", 22, 0) & oImpre.ImpreFormat("Cuota", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA + ITF", 12, 0), "") & Chr$(10)
        sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 72, 60), "-") & Chr$(10)
    End If
        nIntereses = 0
        nCapital = 0
        nCuotaITF = 0
        
        'ARCV 12-02-2007
        nGastos = 0
        nIntGracia = 0
        nGastosComi = 0
        nGastosSegDes = 0
        nGastosEstCta = 0
        nTotCuota = 0 '*** PEAC 20080722
        nGastosPolizaIncendio = 0 '*** PEAC 20090721
        '------
        'WIOR 20120528 ************
        nGastosPolizaMicroseguro = 0
        nGastosPolizaMultiriesgo = 0
        'WIOR FIN *****************
        lnContador = 0
        
        'MAVM 20130209 ***
        If R!nTipoGracia = 6 Then
            Dim nSumMontoGra As Double
            For i = 0 To UBound(MatCalend) - 1
             nSumMontoGra = nSumMontoGra + CDbl(MatCalend(i, 5))
            Next i
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(Format(R!dFecVig + R!nPeriodoGracia, "dd/MM/yyyy")), "dd/mm/yy"), 8, 0) & Space(3)
            sCadImp = sCadImp & "0" & Space(9)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nMonto, 8, 2) & Space(7)
            '->***** LUCV20180601, Agregó condición. según ERS022-2018
             If bVerCalendIntGrac Then
                 sCadImp = sCadImp & oImpre.ImpreFormat(nSumMontoGra, 8, 2) & Space(29)
                 sCadImp = sCadImp & oImpre.ImpreFormat(R!nMonto, 8, 2)
             Else
                sCadImp = sCadImp & Space(35) & oImpre.ImpreFormat(R!nMonto, 8, 2)
             End If
             '<-***** Fin LUCV20180601
            sCadImp = sCadImp & Chr$(10)
        End If
        '***
        
        For i = 0 To UBound(MatCalend) - 1
            
            'ARCV 08-06-2006
            'nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9))
            'ARCV 22-06-2007
               
            '*** PEAC 20090721 - CDbl(MatCalend(i, 14)) poliza de incendio
            'nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 12)) + CDbl(MatCalend(i, 13))
            'nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 12)) + CDbl(MatCalend(i, 13)) + CDbl(MatCalend(i, 14))
            'WIOR 20120528 ******
            'nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 12)) + CDbl(MatCalend(i, 13)) + CDbl(MatCalend(i, 14)) + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))) + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 16)))'RECO20150601 ERS023-2015
            'nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 12)) + CDbl(MatCalend(i, 13)) + CDbl(MatCalend(i, 14)) + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))) + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 16))) + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 17))) 'RECO20150601 ERS023-2015 'LUCV20180601. Coment?
            nCuota = CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 12)) + CDbl(MatCalend(i, 13)) + CDbl(MatCalend(i, 14)) + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))) + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 16))) + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 17))) + CDbl(IIf(MatCalend(i, 18) = "", 0, MatCalend(i, 18))) 'LUCV20180601. Seg?n ERS022-2018
            'WIOR FIN ***********
            '-------
            'DETALLADO
            
            If pbDetallado = True Then
            
                lnContador = lnContador + 1
                           
                ''ARCV 14-02-2007
                ''    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(MatCalend(i, 0)), "ddd, d mmm yyyy"), 17, 0) & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 4, 0) & oImpre.ImpreFormat(nCuota, 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 3)), 9, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 4)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 5)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 9)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 8, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuota * (1 + oITF.gnITFPorcent), 8, 2), "") & Chr$(10)
                'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(MatCalend(i, 0)), "ddd, d mmm yyyy"), 17, 0) & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 4, 0) & oImpre.ImpreFormat(nCuota, 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 3)), 9, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 4)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 5)), 8, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 11)), 10, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 12)), 10, 2) & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 12, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuota * (1 + oITF.gnITFPorcent), 10, 2), "") & Chr$(10)
                ''------
                '**DAOR 20070403***********************************************
                '**Modificado por CAPI 20080105, Se formateo fecha
                
                '*** PEAC 20080722
                'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(Format(MatCalend(i, 0), "dd/MM/yyyy")), "ddd, d mmm yyyy"), 17, 0)
                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(Format(MatCalend(i, 0), "dd/MM/yyyy")), "dd/mm/yy"), 8, 0) '** FECHA
                
                '*** PEAC 20080722
                'sCadImp = sCadImp & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 4, 0)
                'sCadImp = sCadImp & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 4, 0) '** NUM.CUOTA
                'WIOR 20120529
                sCadImp = sCadImp & Space(1)
                sCadImp = sCadImp & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 3, 0) '** NUM.CUOTA
                'WIOR20120529
                sCadImp = sCadImp & oImpre.ImpreFormat(nCuota, 7, 2) '** CUOTA   'WIOR20120529 cambio tama?o de 8 a 7
                sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 3)), 7, 2) '** CAPITAL   'WIOR20120529 cambio tama?o de 8 a 7
                
                '****************** begin
                '->***** LUCV20180601, Agregó condición. según ERS022-2018
                If bVerCalendIntGrac Then
                    sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 4)), 6, 2) '**INT. COMP.
                    sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 5)), 6, 2) '**INT. GRACIA
                Else
                    sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)), 8, 2) '** INT. COMP. + INT. GRACIA
                End If
                '<-***** Fin LUCV20180601
                '****************** end
                
                sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 11)), 6, 2) '** GTOS.Y COMI.
                
                'By Capi 28122007 para insertar nueva columna de gastos excepto a paralelos
                If Not pbCalendParalelo Then
                    '*** PEAC 20080722
                    'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 13)), 8, 2)
                End If
                'End By
                
                sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 12)), 6, 2) '** SEG.DESGRAV.
                
                '*** PEAC 20090721
                'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 14)), 6, 2) '** SEG.POLIZA INCENDIO
                'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 14)), 4, 2)  '** SEG.POLIZA INCENDIO 'JUEZ 20150603 'LUCV20180601. Coment? Seg?n ERS022-2018
                sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 14)) + CDbl(MatCalend(i, 18)), 4, 2) '** SEG.POLIZA INCENDIO 'LUCV20180601
                '***FIN PEAC
                'RECO20150601 ERS023-2015 *******************************************
                'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 17)), 9, 2) '** MULTIRIESGO MYPE
                sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 17)), 3, 2) '** MULTIRIESGO MYPE 'JUEZ 20150603
                'RECO FIN************************************************************
                'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 10, 2) '** SALDO K
                'WIOR 20120528 ************************************************
                If bMicroseguro = True Then
                    sCadImp = sCadImp & Space(3)
                    If bMultiriesgo = True Then
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), 3, 2) 'MICROSEGURO
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 16))), 4, 2) 'MULTIRIESGO
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 7, 2) '** SALDO K
                    Else
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), 3, 2) 'MICROSEGURO
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 10, 2) '** SALDO K
                    End If
                Else
                    If bMultiriesgo = True Then
                        sCadImp = sCadImp & Space(3)
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), 4, 2) 'MULTIRIESGO
                        sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 10, 2) '** SALDO K
                    Else
                     'sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 10, 2) '** SALDO K
                     sCadImp = sCadImp & oImpre.ImpreFormat(CDbl(MatCalend(i, 10)), 8, 2) '** SALDO K 'JUEZ 20150603
                    End If
                End If
                'WIOR FIN *****************************************************
                '*** PEAC 20080722
                'sCadImp = sCadImp & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuota * (1 + oITF.gnITFPorcent), 9, 2), "")
                
                nComiCorresp = Round(nCuota * nFactorCorresp, 2)
                nComiCorresp = IIf(nComiCorresp < nMonMinComi, nMonMinComi, IIf(nComiCorresp > nMonMaxComi, nMonMaxComi, nComiCorresp))
                
                '*** PEAC 20090721
                'sCadImp = sCadImp & oImpre.ImpreFormat(nComiCorresp, 6, 2) & Chr$(10) '** PGO.BCO.NAC.
                sCadImp = sCadImp & Chr$(10)
                
                '**************************************************************
            Else

                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(CDate(MatCalend(i, 0)), "ddd, d mmm yyyy"), 20, 0) & oImpre.ImpreFormat(CInt(MatCalend(i, 1)), 10, 0) & oImpre.ImpreFormat(nCuota, 20, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuota * (1 + oITF.gnITFPorcent), 10, 2), "") & Chr$(10)

            End If
            '*** PEAC 20080827
            'nIntereses = nIntereses + CDbl(MatCalend(i, 4))
            nIntereses = nIntereses + CDbl(MatCalend(i, 4))
            '*** FIN PEAC
            
            nIntereses = CDbl(Format(nIntereses, "#0.00"))
            
            'MAVM
            nIntGracia = nIntGracia + CDbl(MatCalend(i, 5))
            nIntGracia = CDbl(Format(nIntGracia, "#0.00"))
            'MAVM
            
            nCapital = nCapital + CDbl(MatCalend(i, 3))
            nCapital = CDbl(Format(nCapital, "#0.00"))
            nCuotaITF = nCuotaITF + nCuota * (1 + oITF.gnITFPorcent)
            
            '*** PEAC 20080722
            nTotCuota = nTotCuota + nCuota
            nTotCuota = CDbl(Format(nTotCuota, "#0.00"))
            
            'ARCV 12-02-2007
            nGastos = CDbl(Format(nGastos + CDbl(MatCalend(i, 9)), "#0.00"))
            
            '*** PEAC 20080827 - EL INTERES DE GRACIA SE UNION CON EL INT COMPENS.
                'nIntGracia = CDbl(Format(nIntGracia + CDbl(MatCalend(i, 5)), "#0.00"))
            'nIntGracia = 0
            '*** END PEAC
            
            'ARCV 14-02-2007
            nGastosComi = CDbl(Format(nGastosComi + CDbl(MatCalend(i, 11)), "#0.00"))
            
            '*** PEAC 20090721
            'nGastosPolizaIncendio = CDbl(Format(nGastosPolizaIncendio + CDbl(MatCalend(i, 14)), "#0.00")) 'LUCV20180601, Coment?
            nGastosPolizaIncendio = CDbl(Format(nGastosPolizaIncendio + CDbl(MatCalend(i, 14)) + CDbl(MatCalend(i, 18)), "#0.00")) 'LUCV20180601, Agreg? Seg?n ERS022-2018
            nMultMYPE = CDbl(Format(nMultMYPE + CDbl(MatCalend(i, 17)), "#0.00")) 'RECO20150601 ERS023-2015
            'WIOR 20120528 ************************************************
            If bMicroseguro = True Then
                If bMultiriesgo = True Then
                    nGastosPolizaMicroseguro = CDbl(Format(nGastosPolizaMicroseguro + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), "#0.00"))
                    nGastosPolizaMultiriesgo = CDbl(Format(nGastosPolizaMultiriesgo + CDbl(IIf(MatCalend(i, 16) = "", 0, MatCalend(i, 16))), "#0.00"))
                Else
                    nGastosPolizaMicroseguro = CDbl(Format(nGastosPolizaMicroseguro + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), "#0.00"))
                    nGastosPolizaMultiriesgo = 0
                End If
            Else
                If bMultiriesgo = True Then
                    nGastosPolizaMicroseguro = 0
                    nGastosPolizaMultiriesgo = CDbl(Format(nGastosPolizaMultiriesgo + CDbl(IIf(MatCalend(i, 15) = "", 0, MatCalend(i, 15))), "#0.00"))
                Else
                    nGastosPolizaMicroseguro = 0
                    nGastosPolizaMultiriesgo = 0
                End If
            End If
            'WIOR FIN *****************************************************
            nGastosSegDes = CDbl(Format(nGastosSegDes + CDbl(MatCalend(i, 12)), "#0.00"))
            'By Capi 28122007
            nGastosEstCta = CDbl(Format(nGastosEstCta + CDbl(MatCalend(i, 13)), "#0.00"))
            '---------
            
            'If i <> 0 And i Mod 44 = 0 Then
            If i <> 0 And i Mod 39 = 0 Then
            
                lnContador = 0
                
                sCadImp = sCadImp & Chr$(12)
                
                '*** PEAC 20080811
                'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
                Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", "", lsTab, False)
                
                'sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
                If pbDuplicado Then sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10) 'JUEZ 20151229
                'By Capi 05032008
                'MAVM 20140304 ***
                'sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
                'sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
                'sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
                If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
                    sCadImp = sCadImp & lsTab & "Operación AF  : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
                Else
                    sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
                End If
                sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
                sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
                '***
                'sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & Space(1) & "Cuota    : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
                
                'MAVM 20140304 ***
                'sCadImp = sCadImp & lsTab & "Tasa Interes  : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & Space(1) & "Plazo    : " & Trim(Str(R!nPlazo)) & Chr$(10)
                'sCadImp = sCadImp & lsTab & "Mensual" & Chr$(10)
                sCadImp = sCadImp & lsTab & "Plazo         : " & Trim(Str(R!nPlazo)) & Chr$(10)
                '***
                
'                sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Cliente  : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
'                sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
'                sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & Space(1) & "Cuota    : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
'                sCadImp = sCadImp & lsTab & "Interes       : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & Space(1) & "Plazo    : " & Trim(Str(R!nPlazo)) & Chr$(10)
'                sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)

                'MAVM 20140304 ***
                'sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
                If psFecDes = Empty Then
                    sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
                    sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
                Else
                    sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(CDate(psFecDes), "ddd, d mmm yyyy") & Chr$(10)
                    nDiasGracia = CInt(R!nPeriodoGracia) - CInt(DateDiff("D", R!dFecVig, CDate(psFecDes)))
                    sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(IIf(nDiasGracia < 0, 0, nDiasGracia))) & Chr$(10)
                End If
                '***
                
                If pbComodin Then
                    sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Cuota Comodin", 30, 0)
                ElseIf pbMiViv Then
                    sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Mi Vivienda", 30, 0)
                Else
                    sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
                End If
                
                If pbCalendDin Then
                    sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("SI", 30, 0)
                Else
                    sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)
                End If
                sCadImp = sCadImp & Chr$(10)
                sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(nTasaEfectivaAnual, "#0.00") & " %" & Chr$(10)
                'sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual : " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10)
                sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual " & cPagoAnticipado & ": " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10) 'JUEZ 20150505
                
                'RIRO/ANGC 2020
                If campanaReactivate = True Then
                    sCadImp = sCadImp & lsTab & "Comision Cofide   : " & "0.5% " & "Anual" & Chr$(10)
                End If
                
                'MAVM 20140304 ***
                'ALPA 20140619 ***
                If Not (RDesPar.BOF Or RDesPar.EOF) Then
                If RTrim(RDesPar!cConvenioBN) <> "NO" Then
                    sCadImp = sCadImp & lsTab & "CODIGO CREDITO MYPE PAGO BANCO DE LA NACION : " & oImpresora.gPrnNegritaON & sCuentaRefBN & oImpresora.gPrnNegritaOFF & Chr$(10)
                End If
                End If
    
                If RDesPar.RecordCount >= 2 Then
                    sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("DESEMBOLSO", 10, 0) & oImpre.ImpreFormat("ESTADO", 10, 2) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
                    Do While Not RDesPar.EOF
                        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & oImpre.ImpreFormat(RDesPar!nCapital, 10, 2, True) & oImpre.ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
                        RDesPar.MoveNext
                    Loop
                End If

                sCadImp = sCadImp & Chr$(10)
                '***
                
                'CUSCO
                If pbDetallado Then
                    ''ARCV 14-02-2007
                    ''sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gastos", 9, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & Chr$(10)
                    ''sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 106, 90), "-") & Chr$(10)
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 15, 0) & oImpre.ImpreFormat("No Cuota", 10, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gast/Comi.(1)", 15, 0) & oImpre.ImpreFormat("Seg.Desgr.(2)", 15, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & Chr$(10)
                    'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 122, 90), "-") & Chr$(10)
                    ''------
                    
                    '*** PEAC 20080722
                    
'                    '**DAOR 20070403*********************************************
'                    sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 90), "-") & Chr$(10)
'                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 17, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 14, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gast. y", 11, 0) & oImpre.ImpreFormat("Gastos", 11, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & oImpre.ImpreFormat("Pago Bco.", 12, 0) & Chr$(10)
'                    sCadImp = sCadImp & lsTab & Space(17) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(47) & oImpre.ImpreFormat("Comi.(1)", 11, 0) & oImpre.ImpreFormat("Est. Cta.", 11, 0) & oImpre.ImpreFormat("Desgr.(2)", 11, 0) & Space(24) & oImpre.ImpreFormat("Naci?n(3)", 12, 0) & Chr$(10)
'                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 17, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 10, 0) & oImpre.ImpreFormat("Int.Gracia", 12, 0) & oImpre.ImpreFormat("Gast. y", 11, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA+ITF", 12, 0), "") & oImpre.ImpreFormat("Pago Bco.", 12, 0) & Chr$(10)
'                    'sCadImp = sCadImp & lsTab & Space(17) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(44) & oImpre.ImpreFormat("Comi.(1)", 11, 0) & oImpre.ImpreFormat("Desgr.(2)", 11, 0) & Space(24) & oImpre.ImpreFormat("Naci?n(3)", 12, 0) & Chr$(10)
'                    sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 90), "-") & Chr$(10)
'                    '*****************************************************************
                                           
                    'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
                    'WIOR 20120528***
                    sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)
                    'WIOR FIN****
                    
                    '*** PEAC 20090721
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 11, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 14, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 8, 0) & oImpre.ImpreFormat("Gast. y", 9, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & oImpre.ImpreFormat("Pago Bco.", 12, 0) & Chr$(10)
                    'sCadImp = sCadImp & lsTab & Space(11) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(33) & oImpre.ImpreFormat("Comi.(1)", 9, 0) & oImpre.ImpreFormat("Desgr.(2)", 11, 0) & Space(12) & oImpre.ImpreFormat("Naci?n(3)", 12, 0) & Chr$(10)
                    'MAVM 25112009 ***
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 11, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 14, 0) & oImpre.ImpreFormat("Capital", 11, 0) & oImpre.ImpreFormat("Interes", 8, 0) & oImpre.ImpreFormat("Gast. y", 9, 0) & oImpre.ImpreFormat("Seguro", 11, 0) & oImpre.ImpreFormat("Seg.Contra", 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 7, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                    'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
                    'WIOR 20120528***********************************
                        If bMicroseguro = True Then
                            If bMultiriesgo = True Then
                                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Mic-", 6, 0) & oImpre.ImpreFormat("Sg Mul", 8, 0) & oImpre.ImpreFormat("Saldo", 7, 0) & Chr$(10)
                                sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat("Seg.", 6, 0) & oImpre.ImpreFormat("Riesg.", 6, 0) & Space(3) & oImpre.ImpreFormat("Cap.", 8, 0) & Chr$(10)
                            Else
                                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat(" Micro-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                                sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Seg.", 10, 0) & Chr$(10)
                            End If
                        Else
                            If bMultiriesgo = True Then
                                sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Seg Mult-", 10, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                                sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 10, 0) & oImpre.ImpreFormat(" Riesg.", 10, 0) & Chr$(10)
                            Else
                              '->***** LUCV20180601, Agregó condición. según ERS022-2018
                                If bVerCalendIntGrac Then
                                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 8, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.", "Seg."), 7, 0) & oImpre.ImpreFormat("Mult.", 7, 0) & oImpre.ImpreFormat("Saldo Cap", 11, 0) & Chr$(10) 'JUEZ 20150603 'LUCV20180601, Coment? Agreg? condicion
                                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("PYME", 7, 0) & Chr$(10) 'JUEZ 20150724 'LUCV20180601, Agreg? coment?.
                                Else
                                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Gast. y", 8, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.", "Seg."), 7, 0) & oImpre.ImpreFormat("Mult.", 7, 0) & oImpre.ImpreFormat("Saldo Cap", 11, 0) & Chr$(10) 'JUEZ 20150603 'LUCV20180601, Coment? Agreg? condicion
                                    sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("PYME", 7, 0) & Chr$(10) 'JUEZ 20150724 'LUCV20180601, Agreg? coment?.
                                End If
                                '<-***** Fin LUCV20180601
                                'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 10, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.Contra", "Seg.Vehic."), 11, 0) & oImpre.ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
                                'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 10, 0) & oImpre.ImpreFormat("No.", 6, 0) & oImpre.ImpreFormat("Cuota", 11, 0) & oImpre.ImpreFormat("Capital", 9, 0) & oImpre.ImpreFormat("Inter.", 9, 0) & oImpre.ImpreFormat("Inter.", 8, 0) & oImpre.ImpreFormat("Gast. y", 8, 0) & oImpre.ImpreFormat("Seguro", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Seg.", "Seg."), 7, 0) & oImpre.ImpreFormat("Mult.", 7, 0) & oImpre.ImpreFormat("Saldo Cap", 11, 0) & Chr$(10) 'JUEZ 20150603 'LUCV20180601, Comentó Agregó condicion
                                'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 5, 0) & Space(21) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comi.(1)", 10, 0) & oImpre.ImpreFormat("Desgr.(2)", 10, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "incendio", "Multirie"), 11, 0) & Chr$(10)
                                'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Com.(1)", 8, 0) & oImpre.ImpreFormat("Desg(2)", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("MYPE", 7, 0) & Chr$(10) 'JUEZ 20150603
                                'sCadImp = sCadImp & oImpre.ImpreFormat("de Pago", 7, 1) & Space(2) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(20) & oImpre.ImpreFormat("Comp.", 5, 0) & Space(4) & oImpre.ImpreFormat("Gracia", 8, 0) & oImpre.ImpreFormat("Comis. ", 8, 0) & oImpre.ImpreFormat("Desgra.", 8, 0) & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 6, 3) <> "517", "Incend", "Vehic."), 7, 0) & oImpre.ImpreFormat("MYPE", 7, 0) & Chr$(10) 'JUEZ 20150724 'LUCV20180601, Agregó comentó.
                            End If
                        End If
                    'WIOR FIN ***************************************
                    
                    'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("de Pago", 8) & Space(3) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(25) & oImpre.ImpreFormat("(1)", 6) & oImpre.ImpreFormat("Comi.(2)", 9, 0) & oImpre.ImpreFormat("Desgr.(3)", 11, 0) & oImpre.ImpreFormat("incendio", 11, 0) & Chr$(10)
                    '***
                    'sCadImp = sCadImp & lsTab & Space(11) & oImpre.ImpreFormat("Cuota", 7, 0) & Space(25) & oImpre.ImpreFormat("(1)", 6) & oImpre.ImpreFormat("Comi.(2)", 9, 0) & oImpre.ImpreFormat("Desgr.(3)", 11, 0) & oImpre.ImpreFormat("incendio", 11, 0) & Chr$(10)
                    'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
                    '*** FIN PEAC
                    'WIOR 20120528***
                    sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)
                    'WIOR FIN****

                Else
                    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Fecha", 25, 0) & oImpre.ImpreFormat("No Cuota", 22, 0) & oImpre.ImpreFormat("Cuota", 12, 0) & IIf(pbMostrarITF, oImpre.ImpreFormat("CUOTA + ITF", 12, 0), "") & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 72, 60), "-") & Chr$(10)
                End If
            End If
    Next i
    
    'CUSCO
    If pbDetallado Then
        
        '*** PEAC 20080722
        
        'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 100), "-") & Chr$(10)
        'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
        'WIOR 20120528***
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)'RECO20150601 ERS023-2015
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 104)), "-") & Chr$(10) 'RECO20150601 ERS023-2015
        sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10) 'JUEZ 20150603
        'WIOR FIN****
        
        'ARCV 12-02-2007
        'sCadImp = sCadImp & lsTab & "Totales : " & oImpre.ImpreFormat(nCapital, 32, 2) & oImpre.ImpreFormat(nIntereses, 7, 2) & oImpre.ImpreFormat(nCapital + nIntereses, 34, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuotaITF, 10, 2), "") & Chr$(10)
        'ARCV 14-02-2007
        'sCadImp = sCadImp & lsTab & "Totales : " & oImpre.ImpreFormat(nCapital, 30, 2) & oImpre.ImpreFormat(nIntereses, 8, 2) & oImpre.ImpreFormat(nIntGracia, 8, 2) & oImpre.ImpreFormat(nGastos, 8, 2) & oImpre.ImpreFormat(nCapital + nIntereses + nIntGracia + nGastos, 9, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuotaITF, 9, 2), "") & Chr$(10)
        'By Capi 28122007
        
'        If Not pbCalendParalelo Then
'            sCadImp = sCadImp & lsTab & "Totales : " & Space(11) & oImpre.ImpreFormat(nCapital + nIntereses + nIntGracia + nGastos, 9, 2) & oImpre.ImpreFormat(nCapital, 9, 2) & oImpre.ImpreFormat(nIntereses, 8, 2) & oImpre.ImpreFormat(nIntGracia, 8, 2) & oImpre.ImpreFormat(nGastosComi, 8, 2) & oImpre.ImpreFormat(nGastosEstCta, 8, 2) & oImpre.ImpreFormat(nGastosSegDes, 8, 2) & Space(13) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuotaITF, 9, 2), "") & Chr$(10)
'        Else
'            sCadImp = sCadImp & lsTab & "Totales : " & Space(11) & oImpre.ImpreFormat(nCapital + nIntereses + nIntGracia + nGastos, 9, 2) & oImpre.ImpreFormat(nCapital, 9, 2) & oImpre.ImpreFormat(nIntereses, 8, 2) & oImpre.ImpreFormat(nIntGracia, 8, 2) & oImpre.ImpreFormat(nGastosComi, 8, 2) & oImpre.ImpreFormat(nGastosSegDes, 8, 2) & Space(13) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuotaITF, 9, 2), "") & Chr$(10)
'        End If

        '*** PEAC 20090721
        'sCadImp = sCadImp & lsTab & "Totales  :   " & oImpre.ImpreFormat(nTotCuota, 9, 2) & oImpre.ImpreFormat(nCapital, 9, 2) & oImpre.ImpreFormat(nIntereses + nIntGracia, 8, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Totales  :   " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & Chr$(10)
        'WIOR 20120528***********************************
         If bMicroseguro = True Then
            If bMultiriesgo = True Then
                sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & Space(3) & oImpre.ImpreFormat(nGastosPolizaMicroseguro, 3, 2) & Space(1) & oImpre.ImpreFormat(nGastosPolizaMultiriesgo, 4, 2) & Chr$(10)
            Else
                sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nGastosPolizaMicroseguro, 7, 2) & Chr$(10)
            End If
        Else
            If bMultiriesgo = True Then
                '->***** LUCV20180601, Agregó condición. según ERS022-2018
                If bVerCalendIntGrac Then
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nGastosPolizaMultiriesgo, 7, 2) & Chr$(10) 'RECO20150601 ERS023-2015
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nGastosPolizaMultiriesgo, 7, 2) & oImpre.ImpreFormat(nMultMYPE, 9, 2) & Chr$(10) 'RECO20150601 ERS023-2015
                Else
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nGastosComi, 8, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nGastosPolizaMultiriesgo, 7, 2) & Chr$(10)
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nGastosComi, 8, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nGastosPolizaMultiriesgo, 7, 2) & oImpre.ImpreFormat(nMultMYPE, 9, 2) & Chr$(10)
                End If
                '<-***** Fin LUCV20180601
            Else
                'sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & Chr$(10)'RECO20150601 ERS023-2015
                'sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 6, 2) & oImpre.ImpreFormat(nMultMYPE, 9, 2) & Chr$(10) 'RECO20150601 ERS023-2015
                '->***** LUCV20180601, Agregó condición. según ERS022-2018
                If bVerCalendIntGrac Then
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat(nIntereses, 6, 2) & oImpre.ImpreFormat(nIntGracia, 6, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 4, 2) & oImpre.ImpreFormat(nMultMYPE, 3, 2) & Chr$(10) 'JUEZ 20150603
                Else
                    sCadImp = sCadImp & lsTab & "Totales  : " & oImpre.ImpreFormat(nTotCuota, 7, 2) & oImpre.ImpreFormat(nCapital, 8, 2) & oImpre.ImpreFormat((nIntereses + nIntGracia), 8, 2) & oImpre.ImpreFormat(nGastosComi, 6, 2) & oImpre.ImpreFormat(nGastosSegDes, 6, 2) & oImpre.ImpreFormat(nGastosPolizaIncendio, 4, 2) & oImpre.ImpreFormat(nMultMYPE, 3, 2) & Chr$(10)
                End If
                '<-***** Fin LUCV20180601
            End If
        End If
        'WIOR FIN ***************************************
        'sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 137, 100), "-") & Chr$(10)
        'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
        
        '*** FIN PEAC
        'WIOR 20120528***
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)'RECO20150601 ERS023-2015
        'sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 104)), "-") & Chr$(10) 'RECO20150601 ERS023-2015
        sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10) 'JUEZ 20150603
        'WIOR FIN****
    Else
        sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 72, 60), "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & "Totales : " & oImpre.ImpreFormat(nCapital + nIntereses, 40, 2) & IIf(pbMostrarITF, oImpre.ImpreFormat(nCuotaITF, 12, 2), "") & Chr$(10)
        sCadImp = sCadImp & lsTab & String(IIf(pbMostrarITF, 72, 60), "-") & Chr$(10)
    End If
'    sCadImp = sCadImp & Chr$(27) & Chr$(107) & Chr$(2) & Chr$(27) & Chr$(83) & Chr$(1)
'    Set oNCred = New COMNCredito.NCOMCredito
'    sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(oNCred.TasaIntPerDias(R!nTasaInteres, 360) * 100, "#0.00") & "%" & Chr$(10) & Chr$(10)
'    Set oNCred = Nothing
'    sCadImp = sCadImp & Chr$(27) & Chr$(84) & Chr$(1)
'    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(1)

    sCadImp = sCadImp & Chr$(10)
    
    
    If lnContador >= 26 Then
       
        sCadImp = sCadImp & Chr$(12)
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", "", lsTab, False)
        
        sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
        'By Capi 05032008
        'MAVM 20140304 ***
        'sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
        If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
            sCadImp = sCadImp & lsTab & "Operación AF  : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
        Else
            sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
        End If
        sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
        sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
        '***
        'MAVM 20140304 ***
        'sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & Space(1) & "Cuota    : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Tasa Interes  : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & Space(1) & "Plazo    : " & Trim(Str(R!nPlazo)) & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Mensual" & Chr$(10)
        sCadImp = sCadImp & lsTab & "Plazo         : " & Trim(Str(R!nPlazo)) & Chr$(10)
        '***
        'sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
        If psFecDes = Empty Then
            sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
            sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
        Else
            sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & Space(1) & "Vigencia : " & Format(CDate(psFecDes), "ddd, d mmm yyyy") & Chr$(10)
            nDiasGracia = CInt(R!nPeriodoGracia) - CInt(DateDiff("D", R!dFecVig, CDate(psFecDes)))
            sCadImp = sCadImp & lsTab & "Monto         : " & oImpre.ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & Space(1) & "Gracia   : " & Trim(Str(IIf(nDiasGracia < 0, 0, nDiasGracia))) & Chr$(10)
        End If
        '***
        If pbComodin Then
            sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Cuota Comodin", 30, 0)
        ElseIf pbMiViv Then
            sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Mi Vivienda", 30, 0)
        Else
            sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
        End If
        
        If pbCalendDin Then
            sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("SI", 30, 0)
        Else
            sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)
        End If
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(nTasaEfectivaAnual, "#0.00") & " %" & Chr$(10)
        'sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual : " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10)
        sCadImp = sCadImp & lsTab & "Tasa Costo Efectivo Anual " & cPagoAnticipado & ": " & Format(nTasaCostoEfectivoAnual, "#0.00") & " %" & Chr$(10) 'JUEZ 20150505
        
        'RIRO ANGC 2020
        If campanaReactivate = True Then
            sCadImp = sCadImp & lsTab & "Comision Cofide   : " & "0.5% " & "Anual" & Chr$(10)
        End If
        
        'sCadImp = sCadImp & lsTab & String(92, "-") & Chr$(10)
        'WIOR 20120528***
        sCadImp = sCadImp & lsTab & String(IIf(bMicroseguro = True, IIf(bMultiriesgo = True, 103, 102), IIf(bMultiriesgo = True, 102, 92)), "-") & Chr$(10)
        'WIOR FIN****
    End If
    
    'MAVM 201402
    If Not (Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516") Then
        'JUEZ 20150724 ******************************************************************************
        'sCadImp = sCadImp & "(1) Incluye los gastos y comisiones que correspondan seg?n las condiciones pactadas de acuerdo" & Chr$(10)
        'sCadImp = sCadImp & "    al tarifario vigente y a la Hoja Resumen entregada. En el caso de cr?ditos por convenio " & Chr$(10)
        sCadImp = sCadImp & "Los gastos y comisiones serán según las condiciones pactadas de acuerdo al tarifario vigente" & Chr$(10)
        sCadImp = sCadImp & "y a la Hoja Resumen que forma parte del contrato suscrito entre las partes." & Chr$(10) & Chr$(10)
        'sCadImp = sCadImp & "    tambi?n incluye el pago de casilleros conforme al convenio suscrito." & Chr$(10)
        'sCadImp = sCadImp & "(2) Cr?dito con " & IIf(bExisteMasDeUnTitular, "2", "1") & " Titular(es). El seguro de desgravamen corresponde a la tasa de " & Format(nValorGastoSeg, "#0.0000") & "% del " & Chr$(10)
        'sCadImp = sCadImp & "    saldo capital." & Chr$(10) & Chr$(10)
        'END JUEZ ***********************************************************************************
    End If
      
    'By Capi 28022008
'    If RTrim(RDesPar!cConvenioBN) = "NO" Then
'
''        sCadImp = sCadImp & "Importante : Si usted requiere realizar su pago en una Ag. del Bco. de la Nacion debera indicar" & Chr$(10)
''        sCadImp = sCadImp & "             el siguiente N? "
'        If Not (Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516") Then
'
'
'            sCadImp = sCadImp & "IMPORTANTE : Si usted requiere realizar su pago en una Ag. del Bco. de la Naci?n deber? indicar" & Chr$(10)
'
'            'MAVM 20100113 **********
'            sCadImp = sCadImp & "en ventanilla, los siguientes datos importantes:" & Chr$(10)
'            sCadImp = sCadImp & "             ? El c?digo del cr?dito " & Chr$(27) & Chr$(69) & sCuentaRefBN & Chr$(27) & Chr$(70) & Chr$(10)
'            sCadImp = sCadImp & "             ? Para el caso de cr?ditos Mypes - Convenio de Locales Compartidos, indicar el" & Chr$(10)
'            sCadImp = sCadImp & "               c?digo " & Chr$(27) & Chr$(69) & "0540 - Entidad 010 CMAC MAYNAS" & Chr$(27) & Chr$(70) & Chr$(10)
'            sCadImp = sCadImp & "             ? Cuando requiera realizar el pago de sus cuotas a trav?s de las Oficinas del" & Chr$(10)
'
'            sCadImp = sCadImp & "               Banco de la Naci?n, indicar el c?digo " & Chr$(27) & Chr$(69) & "3910 - Cobranza por Cupones." & Chr$(27) & Chr$(70) & Chr$(10)
'
'            sCadImp = sCadImp & "Asimismo el Banco de la Naci?n le cobrar? una comision de acuerdo a su tarifario vigente" & Chr$(10)
'            'MAVM 20100113 **********
'            'WIOR 20130429 ***************************************************************************
'            sCadImp = sCadImp & Chr$(10) & "Si usted requiere realizar su pago en una Ag. del Bco. de Cr?dito del Per? deber? indicar en" & Chr$(10)
'            sCadImp = sCadImp & "ventanilla, los siguientes datos importantes:" & Chr$(10)
'            sCadImp = sCadImp & "             ? El c?digo del cr?dito " & Chr$(27) & Chr$(69) & Mid(psCtaCod, 4, 6) + Mid(psCtaCod, 11, 8) & Chr$(27) & Chr$(70) & Chr$(10)
'            sCadImp = sCadImp & "             ? " & Chr$(27) & Chr$(69) & "Entidad CMAC MAYNAS" & Chr$(27) & Chr$(70) & Chr$(10)
'            sCadImp = sCadImp & "Asimismo el Banco de Cr?dito del Per? le cobrar? una comisi?n de acuerdo a su tarifario vigente." & Chr$(10) & Chr$(10)
'            'WIOR FIN ********************************************************************************
'
'            'Comentado Por MAVM
'            'sCadImp = sCadImp & "             el siguiente N? " 'kkkkkkkkkkkkkk 'asi mismo el Banco el cobrara una comision  por
'                                'el servicio de corresponsalia de acuerdo a sus tarifarios vigentes.
'            '*** PEAC 20080722
'            'sCadImp = sCadImp & oImpresora.gPrnNegritaON & sCuentaRefBN & oImpresora.gPrnNegritaOFF & Chr$(10)
'
'            'Comentado Por MAVM
'            'sCadImp = sCadImp & oImpresora.gPrnBoldON & sCuentaRefBN & oImpresora.gPrnBoldOFF & " asi mismo el Banco le cobrara una comision " & Chr$(10)
'            '*** PEAC 20090720
'            'Comentado Por MAVM
'            'sCadImp = sCadImp & "             por el servicio de corresponsalia de acuerdo a sus tarifarios vigentes." & Chr$(10)
'        End If
'    End If
'    RDesPar.Close
    If Not (RDesPar.BOF Or RDesPar.EOF) Then
    If RTrim(RDesPar!cConvenioBN) = "NO" Then
        If Not (Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516") Then
            sCadImp = sCadImp & Chr$(27) & Chr$(69) & "IMPORTANTE:" & Chr$(27) & Chr$(70) & Chr$(10)
            sCadImp = sCadImp & "Estimado Cliente le informamos que usted puede también realizar su pago a través de los sigui-" & Chr$(10)
            sCadImp = sCadImp & "entes bancos:" & Chr$(10)
            sCadImp = sCadImp & Chr$(27) & Chr$(69) & "1. Banco de la Nación," & Chr$(27) & Chr$(70) & " indicando los siguientes datos:" & Chr$(10)
            sCadImp = sCadImp & "     ? El código del crédito " & Chr$(27) & Chr$(69) & sCuentaRefBN & Chr$(27) & Chr$(70) & " y el código " & Chr$(27) & Chr$(69) & "3910 - Cobranza por Cupones" & Chr$(27) & Chr$(70) & Chr$(10)
            sCadImp = sCadImp & "     ? Si se trata de créditos Mype - Convenio de locales compartidos deberás indicar el código" & Chr$(10)
            sCadImp = sCadImp & "       " & Chr$(27) & Chr$(69) & "0540 - Entidad 010 CMAC MAYNAS." & Chr$(27) & Chr$(70) & Chr$(10)
            sCadImp = sCadImp & Chr$(27) & Chr$(69) & "2. Banco de Crédito del Perú - BCP," & Chr$(27) & Chr$(70) & " indicando los siguientes datos:" & Chr$(10)
            sCadImp = sCadImp & "     ? El código del crédito " & Chr$(27) & Chr$(69) & Mid(psCtaCod, 4, 6) + Mid(psCtaCod, 11, 8) & Chr$(27) & Chr$(70) & Chr$(10)
            sCadImp = sCadImp & "     ? Entidad CMAC MAYNAS" & Chr$(10) & Chr$(10)
        End If
    End If
    End If
    'RDesPar.Close 'COMENTADO X JUEZ 20150803
    
    'ARCV 14-02-2007
'    If Not (Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516") Then
'        '*** PEAC 20090720
'        'sCadImp = sCadImp & "(1)Incluye  consulta central de riesgos, comision administracion de credito y pago de casillero" & Chr$(10)
'        'sCadImp = sCadImp & "(1) Incluye los intereses compensatorios e intereses por per?odo de gracia." & Chr$(10)
'        sCadImp = sCadImp & "(1) Incluye los gastos y comisiones que correspondan segun las condiciones pactadas de acuerdo" & Chr$(10)
'                                'Incluye consulta central de riesgos, comisi?n administraci?n de cr?dito y pago de casillero" & Chr$(10)
'        sCadImp = sCadImp & "    al tarifario vigente. En el caso de creditos por convenio tambien incluye pago de casilleros" & Chr$(10)
'        sCadImp = sCadImp & "    conforme al convenio suscrito." & Chr$(10)
'                                'al tarifario vigente. En el caso de creditos por convenio tambien incluye pago de casilleros
'                                'conforme al convenio suscrito.
'                                'en convenios que correpondan." & Chr$(10)
'        sCadImp = sCadImp & "(2) Credito con " & IIf(bExisteMasDeUnTitular, "2", "1") & " Titular(es). El seguro de desgravamen corresponde a la tasa de " & Format(nValorGastoSeg, "#0.0000") & "% del " & Chr$(10)
'        sCadImp = sCadImp & "    saldo capital + interes de  la cuota." & Chr$(10) & Chr$(10)
'    End If
'    sCadImp = sCadImp & " en convenios que correspondan." & Chr$(10)
'    sCadImp = sCadImp & "(2)Credito con " & IIf(bExisteMasDeUnTitular, "2", "1") & " Titular(es), corresponde la tasa de " & Format(nValorGastoSeg, "#0.0000") & "% del saldo capital + interes de  la" & Chr$(10)
'    sCadImp = sCadImp & "cuota"
'    sCadImp = sCadImp & "(3)Si paga en el Banco de la Nacion adicione a la cuota" & Chr$(10)
    'JUEZ 20150803 ****************************************************************
    'sCadImp = sCadImp & Chr$(27) & Chr$(69) & "NOTA:" & Chr$(27) & Chr$(70) & " Por el pago que efect?e tanto en el Banco de la Naci?n y el BCP, los Bancos le cobra-" & Chr$(10)
    'sCadImp = sCadImp & "r?n una comisi?n por el servicio brindado de acuerdo a su tarifario vigente." & Chr$(10) & Chr$(10)
    'sCadImp = sCadImp & "Todas las operaciones est?n sujetas al pago de ITF " & oITF.gnITFPorcent * 100 & "% por transaccion." & Chr$(10)
    If Not (RDesPar.BOF Or RDesPar.EOF) Then
        If RTrim(RDesPar!cConvenioBN) <> "NO" Then
            sCadImp = sCadImp & Chr$(27) & Chr$(69) & "NOTA:" & Chr$(27) & Chr$(70) & " Por el pago que efectúe en el Banco de la Nación, éste le cobrará una comisión por" & Chr$(10)
            sCadImp = sCadImp & "el servicio brindado de acuerdo a su tarifario vigente." & Chr$(10) & Chr$(10)
        Else
            sCadImp = sCadImp & Chr$(27) & Chr$(69) & "NOTA:" & Chr$(27) & Chr$(70) & " Por el pago que efectúe tanto en el Banco de la Nación y el BCP, los Bancos le cobra-" & Chr$(10)
            sCadImp = sCadImp & "rán una comisión por el servicio brindado de acuerdo a su tarifario vigente." & Chr$(10) & Chr$(10)
        End If
    Else
        sCadImp = sCadImp & Chr$(27) & Chr$(69) & "NOTA:" & Chr$(27) & Chr$(70) & " Por el pago que efectúe tanto en el Banco de la Nación y el BCP, los Bancos le cobra-" & Chr$(10)
        sCadImp = sCadImp & "rán una comisión por el servicio brindado de acuerdo a su tarifario vigente." & Chr$(10) & Chr$(10)
    End If
    sCadImp = sCadImp & "Todas las operaciones están sujetas al pago de ITF " & oITF.gnITFPorcent * 100 & "% por transacción." & Chr$(10)
    RDesPar.Close
    'END JUEZ **********************************************************************
        
    '-------------
    
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        'sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
        sCadImp = sCadImp & lsTab & "Si no se ha Realizado su Descuento por Planilla" & Chr$(10)
        sCadImp = sCadImp & lsTab & "Acerquese a nuestras Agencias a Cancelar y Evite el cobro de Mora" & Chr$(10)
    End If
    If CInt(Mid(psCtaCod, 6, 3)) = gColHipoMiVivienda Or CInt(Mid(psCtaCod, 6, 3)) = gColHipoCaja Then
        'sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
        'MAVM 25112009 ***
        'sCadImp = sCadImp & "NOTA : Todas los Cuotas incluyen Gastos, comisiones y seguros." & Chr$(10)
    End If
    
    'function por credicom para verificar si tiene o no convenion con los gastos
    Set oCredDoc = New COMDCredito.DCOMCredDoc
    bEstadoConvenio = oCredDoc.VerificaConvenioEmp(psCtaCod)
    Set oCredDoc = Nothing
    
    If bEstadoConvenio = True Then
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & " Se debe indicar que los siguientes costos por comision de servicios (COM) y Aportes  de" & Chr$(10)
        sCadImp = sCadImp & " los socios (A. Socios) no forman parte de los interes ni gastos que originan el crédito" & Chr$(10)
        sCadImp = sCadImp & " del presente cronograma en tal sentido la CAJA no tiene ninguna  responsabilidad por el" & Chr$(10)
        sCadImp = sCadImp & " uso o destino  de dichos conceptos." & Chr$(10)
    End If
        
    '**DAOR 20070504*****************************************************
    '    sCadImp = sCadImp & "En caso de reclamos y/o denuncias por las operaciones y servicios que realice, el cliente podr?" & Chr$(10)
    '    sCadImp = sCadImp & "recurrir a las distintas instancias, tales como:" & Chr$(10)
    '    sCadImp = sCadImp & "(a) Red de Agencias de la Caja Maynas. " & Chr$(10)
    '    sCadImp = sCadImp & "(b) Plataforma de Atenci?n al Usuario de la Superintendencia de Banca Seguros y AFP" & Chr$(10)
    '    sCadImp = sCadImp & "(c) Indecopi " & Chr$(10)
    '*******************************************************************
    
'    'MAVM 20110914 ***
'    sCadImp = sCadImp & "1.- EL CLIENTE tiene la facultad de efectuar pagos anticipados o prepagos en forma total o parcial," & Chr$(10)
'    sCadImp = sCadImp & "    con la consiguiente reducci?n de los intereses al d?a de pago, deduci?ndose as? mismo las " & Chr$(10)
'    sCadImp = sCadImp & "    comisiones y gastos derivados de las condiciones contractuales pactadas entre las partes, sin" & Chr$(10)
'    sCadImp = sCadImp & "    que le sean aplicables comisiones, gastos o penalidades de alg?n tipo o cobros de naturaleza" & Chr$(10)
'    sCadImp = sCadImp & "    o efecto similar." & Chr$(10)
'
'    sCadImp = sCadImp & "2.- Si producto de dolo o culpa debidamente acreditados EL CLIENTE es inducido a error por parte de" & Chr$(10)
'    sCadImp = sCadImp & "    LA CAJA y como consecuencia de ello realiza un pago en exceso, dicho importe ser? recuperable y" & Chr$(10)
'    sCadImp = sCadImp & "    devengar? hasta su devoluci?n el m?ximo de la suma por concepto de interes compensatorio y" & Chr$(10)
'    sCadImp = sCadImp & "    moratorio que se hayan pactado para esta operaci?n o en su defecto, el inter?s legal." & Chr$(10)
'
'    'MAVM 20130209 ***
'    If R!nTipoGracia = 1 Or R!nTipoGracia = 6 Then
'        sCadImp = sCadImp & "3.- El metodo de calculo de interes de gracia elegido por el cliente es: " & R!cTipoGracia & Chr$(10)
'    End If
'    '***
'    '***
    
    Set oITF = Nothing
    'sCadImp = sCadImp & Chr$(12)
    ImprimePlandePagos = sCadImp
End Function

Public Function ImprimeComprobanteDesembolso(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, ByVal psNomCmac As String) As String
    
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim oCred As COMDCredito.DCOMCredito
Dim oCalend As COMDCredito.DCOMCalendario
Dim R As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RDesemb As ADODB.Recordset
Dim sCadImp As String
Dim sPersNombreUser As String

    Set oCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oCredDoc.RecuperaDatosDocDesembolso(psCtaCod)
    sPersNombreUser = oCredDoc.ObtieneNombrePersonaUser(psCodUsu)

    Set oCredDoc = Nothing
    
    Set oCred = New COMDCredito.DCOMCredito
    Set REstadosCred = oCred.RecuperaEstadosdelCredito(psCtaCod, True, True, True)
    Set oCred = Nothing
    Set oCalend = New COMDCredito.DCOMCalendario
    
    Set oCalend = New COMDCredito.DCOMCalendario
    Set RDesemb = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
   
    
    sCadImp = ""
    If R.EOF Then Exit Function
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "COMPROBANTE DE DESEMBOLSO", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & "Operación AF        : " & oImpre.ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "Credito             : " & oImpre.ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & "Fecha de Aprobacion : " & oImpre.ImpreFormat(Format(R!dFecApr, "dd/mm/yyyy"), 20, 0) & "Codigo de Cliente : " & R!cPersCod & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & oImpre.ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & oCad.PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 20, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("CUOTAS", 18, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        'MAVM 20121220 ***
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nTipoGracia = 1, REstadosCred!nCuotas + 1, REstadosCred!nCuotas), 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc 'COMENTADO POR WIOR 20130223
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nTipoGracia = 1, REstadosCred!nCuotas + IIf(REstadosCred!nPeriodoGracia > 0, 1, 0), REstadosCred!nCuotas), 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc 'WIOR 20130223 AGREGO IIf(REstadosCred!nPeriodoGracia > 0, 1, 0)
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nTipoGracia = 1, REstadosCred!nCuotas, REstadosCred!nCuotas), 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        '***
        REstadosCred.MoveNext
    Loop
    REstadosCred.Close
    Set REstadosCred = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APODERADO         : " & oCad.PstaNombre(R!cApoderado) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA            : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO  : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INTERES           : " & Format(R!nTasaInteres, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TIPO DE DESEMBOLSO: " & R!cTipoDesemb & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(25) & "DESEMBOLSOS EFECTUADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("F. PACTADA", 15, 0) & oImpre.ImpreFormat("FECHA", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("GASTOS", 15, 0) & oImpre.ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Format(RDesemb!dPago, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Trim(Str(RDesemb!nCuota)), 5, 0) & oImpre.ImpreFormat(RDesemb!nCapital, 15, 2) & oImpre.ImpreFormat(RDesemb!nGasto, 10, 2) & Space(10) & IIf(RDesemb!ncoloccalendestado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & "________________________________   ______________________________    _________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(29, "_") & Space(3) & String(29, "_") & Space(3) & String(29, "_") & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & ocad.PstaNombre(R!cPersNombre) & Space(5) & ocad.PstaNombre(R!cAnalista) & Space(15) & "CAJA" & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & oCad.PstaNombre(R!cPersNombre) & Space(5) & oCad.PstaNombre(sPersNombreUser) & Space(19) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(oCad.PstaNombre(R!cPersNombre), 29) & Space(5) & oImpre.ImpreFormat(oCad.PstaNombre(sPersNombreUser), 29) & Space(19) & lnSaltoLinDoc
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    ImprimeComprobanteDesembolso = sCadImp
    
End Function

'LUCV20160621, Según ERS004-2016 ->**********
Public Function ImprimeComprobanteVigencia(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, ByVal psNomCmac As String, Optional ByVal pnRefina As Integer = -1) As String
    
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim oCred As COMDCredito.DCOMCredito
Dim oCalend As COMDCredito.DCOMCalendario
Dim R As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RDesemb As ADODB.Recordset
Dim sCadImp As String
Dim sPersNombreUser As String

    Set oCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oCredDoc.RecuperaDatosDocDesembolso(psCtaCod, pnRefina)
    sPersNombreUser = oCredDoc.ObtieneNombrePersonaUser(psCodUsu)

    Set oCredDoc = Nothing
    
    Set oCred = New COMDCredito.DCOMCredito
    Set REstadosCred = oCred.RecuperaEstadosdelCredito(psCtaCod, True, True, True)
    Set oCred = Nothing
    Set oCalend = New COMDCredito.DCOMCalendario
    
    Set oCalend = New COMDCredito.DCOMCalendario
    Set RDesemb = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
       
    sCadImp = ""
    If R.EOF Then Exit Function
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "COMPROBANTE DE VIGENCIA", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & "Operación AF        : " & oImpre.ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "Credito             : " & oImpre.ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & "Fecha de Aprobacion : " & oImpre.ImpreFormat(Format(R!dFecApr, "dd/mm/yyyy"), 20, 0) & "Codigo de Cliente : " & R!cPersCod & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & oImpre.ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & oCad.PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 20, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("CUOTAS", 18, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 10, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nTipoGracia = 1, REstadosCred!nCuotas, REstadosCred!nCuotas), 15, 0) & oImpre.ImpreFormat(REstadosCred!nPlazo, 15, 0) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    REstadosCred.Close
    Set REstadosCred = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APODERADO         : " & oCad.PstaNombre(R!cApoderado) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA            : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO  : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INTERES           : " & Format(R!nTasaInteres, "#0.00") & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & "TIPO DE DESEMBOLSO: " & R!cTipoDesemb & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(25) & "VIGENCIA DE CREDITOS REFINANCIADOS - EFECTUADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("F. PACTADA", 15, 0) & oImpre.ImpreFormat("FECHA", 15, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("GASTOS", 15, 0) & oImpre.ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Format(RDesemb!dPago, "dd/mm/yyyy"), 15, 0) & oImpre.ImpreFormat(Trim(Str(RDesemb!nCuota)), 5, 0) & oImpre.ImpreFormat(RDesemb!nCapital, 15, 2) & oImpre.ImpreFormat(RDesemb!nGasto, 10, 2) & Space(10) & IIf(RDesemb!ncoloccalendestado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & "________________________________   ______________________________    _________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(29, "_") & Space(3) & String(29, "_") & Space(3) & String(29, "_") & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & ocad.PstaNombre(R!cPersNombre) & Space(5) & ocad.PstaNombre(R!cAnalista) & Space(15) & "CAJA" & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & oCad.PstaNombre(R!cPersNombre) & Space(5) & oCad.PstaNombre(sPersNombreUser) & Space(19) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(oCad.PstaNombre(R!cPersNombre), 29) & Space(5) & oImpre.ImpreFormat(oCad.PstaNombre(sPersNombreUser), 29) & Space(19) & lnSaltoLinDoc
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    ImprimeComprobanteVigencia = sCadImp
End Function
'Fin LUCV20160621 <-**********


Function EmiteBoletaAhorroAbonoRetiro(ByVal sCuenta As String, ByVal pnMonto As Double, ByVal nSaldoDisp As Double, _
    ByVal nSaldoCnt As Double, ByVal nIntGanado As Double, ByVal pnTipoOpeAho As Integer, _
    ByVal pdFecSis As Date, ByVal sNomAge As String, ByVal sCodUser As String, Optional ByVal psCodCMAC As String, _
    Optional ByVal pnITF As Double = 0, Optional ByVal pbSaldoInt As Boolean = True, Optional ByVal psCtaCodCred As String = "", Optional ByVal psNombreTitularEco As String = "") As String
    'ALPA 20130103  se agregó psCtaCodCred
    'BRGO 20120502 Se agregó pbSaldoInt
Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String, sNroDoc As String
Dim sMsgOpe  As String
Dim sCadOpe As String

If pnTipoOpeAho = 1 Then
    sModDep = "Deposito Efectivo"
    sTipApe = "DEPOSITO AHORROS"
    sMsgOpe = "Deposito OP"
    sCadOpe = gAhoDepEfec
Else
    sModDep = "Retiro Efectivo"
    sTipApe = "RETIRO AHORROS"
    sMsgOpe = "Retiro OP"
    sCadOpe = gAhoRetEfec
    
    If bAplicaComision Then
    
        'RIRO 20200706 ****************
        Dim sMensajeReactiva As String
        Dim oGen As COMDConstSistema.DCOMGeneral
        Set oGen = New COMDConstSistema.DCOMGeneral
        sMensajeReactiva = oGen.LeeConstSistema("725")
        If Not Len(Trim(sMensajeReactiva)) > 0 Then
            sMensajeReactiva = "Retiro Comisión Reactiva"
        End If
        'END RIRO *********************
        'sMsgOpe = "Retiro OP" RIRO 20200706 Comentado
        
        sMsgOpe = sMensajeReactiva
        
    End If
    
End If

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
sNomTit = oImpre.ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

EmiteBoletaAhorroAbonoRetiro = ImprimeBoletaAhorro(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCadOpe, _
    Trim(pnMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, pbSaldoInt _
    , , , , , , , , pdFecSis, sNomAge, sCodUser, "", psCodCMAC, pnITF, psCtaCodCred, psNombreTitularEco)
'ALPA 20121226 se agregó psNombreTitularEco

End Function


'**************************************************************
'Arma Recibo de Apertura de Cuentas Ahorros
'***************************************************************
Function EmiteBoletaAhorro(ByVal sCuenta As String, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
                        ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal psCodUser As String, Optional ByVal psCodCMAC As String, _
                        Optional ByVal psLpt As String, Optional ByVal pnITF As Double) As String
Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String, sNroDoc As String

sModDep = "Depósito Efectivo"
sTipApe = "APERTURA AHORROS"

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
'Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento   'NCapMovimientos
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales   'NCapMantenimiento
sNomTit = oImpre.ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

EmiteBoletaAhorro = ImprimeBoletaAhorro(sTipApe, oImpre.ImpreCarEsp(sModDep), sCodOpe, Trim(nSaldoCnt), sNomTit, sCuenta, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , pdFecSis, Trim(psNomAge), psCodUser, psLpt, psCodCMAC, pnITF)

End Function

'Se volvio Cadena para pasar al Cliente la CADENA

Public Function ImprimeBoletaAhorro(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional ByVal psCodCMAC As String, Optional ByVal pnMontoIFT As Double, Optional ByVal psCtaCodCred As String = "", Optional ByVal psNombreTitularEco As String = "") As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39
Dim sText As String


''  quitar este comentario

        Dim clsGen As COMDConstSistema.DCOMGeneral
        Set clsGen = New COMDConstSistema.DCOMGeneral
 

ETIQ:

'On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

'nFicSal = FreeFile
'If sLpt <> "" Then
'    Open sLpt For Output As nFicSal
'End If

 ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(64)

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(18)  ' cancela condensada
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(72)  ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)

'Print #nFicSal, Chr$(10);
If psCodCMAC = "102" Then  ' CAJA LIMA
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & lsNegritaOff   'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))

                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        'IMPRESION DE ITF
        
   
        If pbSaldoInt Then
            'If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
            'Else
            '    Print #nFicSal, ""
            '    Print #nFicSal, ""
            '    pbSaldoInt = False
            'End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)  ''& Space(29 + sSep + sAux)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & Chr$(10)
        End If
        'Dim clsGen As COMDConstSistema.DCOMGeneral
        'Set clsGen = New COMDConstSistema.DCOMGeneral
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsMensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux

Else   '' Caja Trujillo
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Space(19 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff  'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        'impresion de ITF
        If pnMontoIFT > 0 Then 'JUEZ 20130930 Para imprimir sólo cuando existe ITF
            sText = "ITF Efectivo"
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Space(-1 + 15) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Chr$(10)
        End If
        If pbSaldoInt Then
            'If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
            'Else
            '    Print #nFicSal, ""
            '    Print #nFicSal, ""
            '    pbSaldoInt = False
            'End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Chr$(10)
        End If
'        Dim clsGen As COMDConstSistema.DCOMGeneral
'        Set clsGen = New COMDConstSistema.DCOMGeneral
        'ALPA 20130103************************************
        If Mid(Trim(psCtaCodCred), 6, 3) = "517" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp("Cliente : " & Left(Trim(psNombreTitularEco) & Space(40), 29)) & Space(14) & oImpre.ImpreCarEsp("Cliente : " & Trim(psNombreTitularEco)) & Chr$(10)
        End If
        '*************************************************
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsMensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux
        
End If
'''- Fin de Liquidaciones

'   If sLpt <> "" Then
'       Print #nFicSal, ImprimeBoletaAhorro
'       Print #nFicSal, Chr$(12)
'       Close nFicSal
'   End If

Exit Function

'Error:
'    Close nFicSal
    'If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
    '    GoTo ETIQ
    'End If
End Function


Public Function ImprimePagoLote(ByVal MatPagos As Variant, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
    ByVal psNomAge As String, ByVal psFecha As String, ByVal psCodUsu As String, ByVal psInstitucion As String, _
    ByVal psMoneda As String, ByVal psNomCmac As String, Optional pbCobrarMora As Boolean = True) As String
'NAGL 20190719 Agregó Optional pbCobrarMora As Boolean = True Según ERS036-2018
Dim i As Integer
Dim sCadImp As String
Dim oNCredito As COMNCredito.NCOMCredito
Dim nMontoPag As Double
Dim nAcumTot As Double
Dim nAcumCapPag As Double
Dim nAcumIntComPag As Double
Dim nAcumIntMorPag As Double
Dim nAcumGastosPag As Double
Dim nAcumITFPag As Double

'Agregado por LMMD
Dim nProxCuota As String
Dim nProFecVenc As String
Dim rs As ADODB.Recordset

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Reporte de Pagos de Prestamos Personales", psNomCmac, lsTab, True)
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & psInstitucion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & psMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(165, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("CREDITO", 25, 0) & oImpre.ImpreFormat("TITULAR", 30, 0) & oImpre.ImpreFormat("CUOTA QUE PAGO", 20, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("CAPITAL", 15, 0) & oImpre.ImpreFormat("INTERES", 12, 0) & oImpre.ImpreFormat("MORA", 8, 0) & oImpre.ImpreFormat("GASTOS", 8, 0) & oImpre.ImpreFormat("I.T.F.", 6, 2) & Space(2) & oImpre.ImpreFormat("P.Cuota", 8, 0) & oImpre.ImpreFormat("F.PROX", 12) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(165, "-") & lnSaltoLinDoc
    
    Set oNCredito = New COMNCredito.NCOMCredito
    nAcumTot = 0
    nAcumCapPag = 0
    nAcumIntComPag = 0
    nAcumIntMorPag = 0
    nAcumGastosPag = 0
    For i = 0 To UBound(MatPagos) - 1
        nMontoPag = CDbl(MatPagos(i, 1)) + CDbl(MatPagos(i, 9))
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(MatPagos(i, 0), 21) & oImpre.ImpreFormat(MatPagos(i, 3), 30) & oImpre.ImpreFormat(IIf(oNCredito.MatrizCuotasPagadas(MatCalendDistrib(i)) = 0, ObtenerCuotaPagada(MatPagos(i, 0)), oNCredito.MatrizCuotasPagadas(MatCalendDistrib(i))), 10) _
        & oImpre.ImpreFormat(nMontoPag, 12) & oImpre.ImpreFormat(oNCredito.MatrizCapitalPagado(MatCalendDistrib(i)), 12) & oImpre.ImpreFormat(oNCredito.MatrizIntCompPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(i)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(i)), 10) & oImpre.ImpreFormat(oNCredito.MatrizIntMorPagado(MatCalendDistrib(i), pbCobrarMora), 8) & oImpre.ImpreFormat(oNCredito.MatrizGastoPag(MatCalendDistrib(i)), 6) & oImpre.ImpreFormat(CDbl(MatPagos(i, 9)), 6, 2)
        'oImpre.ImpreFormat(oNCredito.MatrizIntMorPagado(MatCalendDistrib(i), pbCobrarMora), 8) 'NAGL 20190724 Agregó pbCobrarMora Según ERS036-2018
        nProxCuota = ""
        nProFecVenc = ""
         Set rs = ObtenerProxCuota(MatPagos(i, 0))
         If Not rs.EOF And Not rs.BOF Then
            nProxCuota = rs!nCuota
            nProFecVenc = Format(rs!dVenc, "dd/mm/yyyy")
         End If
         Set rs = Nothing
         sCadImp = sCadImp & Space(2) & oImpre.ImpreFormat(nProxCuota, 3, 0) & oImpre.ImpreFormat(nProFecVenc, 10) & lnSaltoLinDoc
         
         nAcumTot = nAcumTot + nMontoPag
         nAcumCapPag = nAcumCapPag + oNCredito.MatrizCapitalPagado(MatCalendDistrib(i))
         nAcumIntComPag = nAcumIntComPag + oNCredito.MatrizIntCompPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(i)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(i)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(i))
         nAcumIntMorPag = nAcumIntMorPag + oNCredito.MatrizIntMorPagado(MatCalendDistrib(i), pbCobrarMora) 'NAGL 20190724 Según ERS036-2018
         nAcumGastosPag = nAcumGastosPag + oNCredito.MatrizGastoPag(MatCalendDistrib(i))
         nAcumITFPag = nAcumITFPag + CDbl(MatPagos(i, 9))
    Next i
    Set oNCredito = Nothing
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " PAGO TOTAL     : " & oImpre.ImpreFormat(nAcumTot, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " CAPITAL PAGADO : " & oImpre.ImpreFormat(nAcumCapPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " INTERES PAGADO : " & oImpre.ImpreFormat(nAcumIntComPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " MORA PAGADA    : " & oImpre.ImpreFormat(nAcumIntMorPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " GASTO PAGADO   : " & oImpre.ImpreFormat(nAcumGastosPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " ITF PAGADO     : " & oImpre.ImpreFormat(nAcumITFPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    
    ImprimePagoLote = sCadImp
    
End Function

Function ObtenerCuotaPagada(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    
    sSql = "Select Max(MCD.nNroCuota) as nCuota"
    sSql = sSql & " From MovColDet MCD"
    sSql = sSql & " Where MCD.nMovNro=(Select Max(M.nMovNro)"
    sSql = sSql & "           From MovCol MC"
    sSql = sSql & "                Inner Join Mov M on M.nMovNro=MC.nMovNro"
    sSql = sSql & "           Where MC.cCtaCod='" & psCtaCod & "' and M.nMovFlag=0) and"
    sSql = sSql & " MCD.cOpeCod not Like '107[123456789]%'" & " and MCD.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    
    If Not rs.EOF And Not rs.BOF Then
       ObtenerCuotaPagada = IIf(IsNull(rs!nCuota), 1, rs!nCuota)
    End If
    Set rs = Nothing
End Function

Function ObtenerProxCuota(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    'Set oConec = New COMConecta.DCOMConecta
    'oConec.AbreConexion
    
    sSql = "Select CC.nCuota,CC.dVenc"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & "       Inner Join ColocacCred C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1 and CC.nColocCalendEstado=0 and "
    sSql = sSql & "       CC.nCuota=(Select Min(CC.nCuota)"
    sSql = sSql & "             From ColocCalendario CC"
    sSql = sSql & "               Inner Join ColocacCred C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen"
    sSql = sSql & "         Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1 and CC.nColocCalendEstado=0)             "
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerProxCuota = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ImpreRepor_HistorialCliente(ByVal pscPerscod As String, ByVal psNomAge As String, _
                                ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim i As Integer
    Dim sCadImp As String
    Dim nTotal As Integer
    Dim rs As ADODB.Recordset
    Dim oDCredDoc As COMDCredito.DCOMCredDoc
    Dim nMontoSoles As Double
    Dim nMontoDolares As Double
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "HISTORIAL DEL CLIENTE", psNomCmac)
    
    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set rs = oDCredDoc.Repor_HistorialCliente(pscPerscod)
    Set oDCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
    '    RaiseEvent IniciaBarra(nTotal)
    Else
        ImpreRepor_HistorialCliente = ""
        Exit Function
    End If
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & " AGENCIA: " & psNomAge & Chr(10)
    sCadImp = sCadImp & " CLIENTE: " & rs!cNomCli & Chr(10)
    sCadImp = sCadImp & " DOC.IDENTIDAD: " & rs!cPersIdNro & Chr(10)
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    sCadImp = sCadImp & Chr(10)


    'ARCV 29-01-2007
    'sCadImp = sCadImp & oImpre.ImpreFormat("CREDITO", 18) & oImpre.ImpreFormat("ESTADO", 9) & oImpre.ImpreFormat("PARTIC.", 9) & oImpre.ImpreFormat("DESEMBOLSO", 10) & oImpre.ImpreFormat("SALDO.CAP", 8) & oImpre.ImpreFormat("NOTA", 6) & oImpre.ImpreFormat("FEC.APROB", 10) & oImpre.ImpreFormat("F.CUL/CAN", 10) & oImpre.ImpreFormat("F.DESEM", 8) & oImpre.ImpreFormat("ANA", 3) & oImpre.ImpreFormat("PROD.", 6) & oImpre.ImpreFormat("INSTITUCION", 15) & oImpre.ImpreFormat("MONEDA", 8) & oImpre.ImpreFormat("DIAS ATRASO", 12) & Chr(10)
    sCadImp = sCadImp & oImpre.ImpreFormat("CREDITO", 18) & oImpre.ImpreFormat("ESTADO", 18) & oImpre.ImpreFormat("PARTIC.", 9) & oImpre.ImpreFormat("DESEMBOLSO", 10) & oImpre.ImpreFormat("SALDO.CAP", 8) & oImpre.ImpreFormat("NOTA", 6) & oImpre.ImpreFormat("F.CUL/CAN", 10) & oImpre.ImpreFormat("F.DESEM", 8) & oImpre.ImpreFormat("ANA", 3) & oImpre.ImpreFormat("MONEDA", 8) & oImpre.ImpreFormat("DIAS ATRASO", 12) & Chr(10)
    '------------------
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    
    rs.MoveFirst

    nMontoSoles = 0
    nMontoDolares = 0
    Do Until rs.EOF
        'ARCV 29-01-2007
        'sCadImp = sCadImp & oImpre.ImpreFormat(rs!cCtaCod, 18) & oImpre.ImpreFormat(rs!EstadoCredito, 8) & oImpre.ImpreFormat(rs!Partic, 10) & oImpre.ImpreFormat(rs!nMontoCol, 8, 2) & oImpre.ImpreFormat(rs!nSaldo, 8, 2) & oImpre.ImpreFormat(rs!nNota, 4, 0) & Space(2) & oImpre.ImpreFormat(Format(rs!dPrdEstado, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!dVigencia, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(IIf(IsNull(rs!cUser), "    ", rs!cUser), 4) & oImpre.ImpreFormat(rs!nTipoCredCod, 5) & oImpre.ImpreFormat(IIf(IsNull(rs!cNomIntitucion), "SIN CONVENIO", rs!cNomIntitucion), 15) & oImpre.ImpreFormat(rs!Moneda, 10) & oImpre.ImpreFormat(rs!nDiasAtrasoAcum, 4, 0) & lnSaltoLinDoc
        '**DAOR 20071121*********************************
        'sCadImp = sCadImp & oImpre.ImpreFormat(rs!cCtaCod, 18) & oImpre.ImpreFormat(rs!EstadoCredito, 18) & oImpre.ImpreFormat(rs!Partic, 10) & oImpre.ImpreFormat(rs!nMontoCol, 8, 2) & oImpre.ImpreFormat(rs!nSaldo, 8, 2) & oImpre.ImpreFormat(rs!nNota, 4, 0) & Space(2) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!dVigencia, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(IIf(IsNull(rs!cUser), "    ", rs!cUser), 4) & oImpre.ImpreFormat(rs!Moneda, 10) & oImpre.ImpreFormat(rs!nDiasAtrasoAcum, 4, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat(rs!cCtaCod, 18) & oImpre.ImpreFormat(rs!EstadoCredito, 18) & oImpre.ImpreFormat(rs!Partic, 10) & oImpre.ImpreFormat(rs!nMontoCol, 8, 2) & oImpre.ImpreFormat(rs!nSaldo, 8, 2) & oImpre.ImpreFormat(rs!nNota, 4, 0) & Space(2) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!dDesembolso, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(IIf(IsNull(rs!cUser), "    ", rs!cUser), 4) & oImpre.ImpreFormat(rs!Moneda, 10) & oImpre.ImpreFormat(rs!nDiasAtrasoAcum, 4, 0) & lnSaltoLinDoc
        '************************************************
        '----------------------------------
        If Mid(rs!cCtaCod, 9, 1) = "1" Then
            nMontoSoles = nMontoSoles + rs!nSaldo
        Else
            nMontoDolares = nMontoDolares + rs!nSaldo
        End If
        rs.MoveNext
    Loop
    
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Chr(10)
   
    sCadImp = sCadImp & Space(10) & "DEUDA CAPITAL EN SOLES: " & Format(nMontoSoles, "#0.00") & Space(5) & "DEUDA CAPITAL EN DOLARES: " & Format(nMontoDolares, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc

   ImpreRepor_HistorialCliente = sCadImp
End Function

'FRHU RQ13790 20140106-  Se agrego el parametro pbPreferencial
Public Function GeneraImpresionPosicionCliente(ByVal pbPreferencial As Boolean, ByVal psNomCmac As String, ByVal pdFecSis As Date, _
                                            ByVal psNomAge As String, ByVal psPersCod As String, _
                                            ByVal psNomPers As String, ByVal psDocNat As String, _
                                            ByVal psDocJur As String, ByVal pnNumCreditos As Integer, _
                                            ByVal pMatCreditos As Variant, ByVal pnNumAhorros As Integer, _
                                            ByVal pMatAhorros As Variant, ByVal psSolesAho As String, _
                                            ByVal psDolaresAho As String, ByVal pnNumPrendario As Integer, _
                                            ByVal pMatPrendario As Variant, ByVal psSolesPig As String, _
                                            ByVal psDolaresPig As String, ByVal pnNumJudicial As Integer, _
                                            ByVal pMatJudicial As Variant, ByVal pnNumCartaFianza As Integer, _
                                            ByVal pMatCartaFianza As Variant, ByVal psComentario As String, _
                                            ByVal pnSaldosAho As Integer, ByVal pMatSistFinac As Variant, Optional psUser As String = "")


Dim i As Integer
Dim ContLineas As Integer
Dim R As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredito

Dim sImpresion As String
Dim sFechaHora As String
Dim oFecha As COMDConstSistema.DCOMGeneral

Dim J As Integer
Dim psCalifiFinal As String
psCalifiFinal = ""
On Error GoTo ErrorGeneraImpresionPosicionCliente

    Set oFecha = New COMDConstSistema.DCOMGeneral
    sFechaHora = Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss")
    Set oFecha = Nothing
    
    sImpresion = ""
    sImpresion = lnSaltoLinDoc
    
    sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & psNomAge & Space(110) & "Usuario: " & psUser & lnSaltoLinDoc & lnSaltoLinDoc
    sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
    'FRHU RQ13790 20140106
    If pbPreferencial = True Then
    sImpresion = sImpresion & Space(60) & "     CLIENTE PREFERENCIAL     " & lnSaltoLinDoc & lnSaltoLinDoc
    End If
    'FIN FRHU RQ13790
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    
    sImpresion = sImpresion & Space(5) & "SECCION DE CREDITOS : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    'sImpresion = sImpresion & Space(5) & String(175, "-") & lnSaltoLinDoc 'Segmentacion Riesgo por cliente JOEP20201113
    
    'INICIO ORCR-20140913*********
    'sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Tipo", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
    'sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Analista", 7) & oImpre.ImpreFormat("A. Inicial", 7) & oImpre.ImpreFormat("Nota", 5) & oImpre.ImpreFormat("Cal.", 3) & oImpre.ImpreFormat("Monto", 10) & oImpre.ImpreFormat("Saldo Cap", 10) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 10) & oImpre.ImpreFormat("Tipo", 12) & oImpre.ImpreFormat("Agencia", 8) & oImpre.ImpreFormat("Cuenta", 18) & oImpre.ImpreFormat("Estado", 9)
    sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 9) & oImpre.ImpreFormat("A.", 5) & oImpre.ImpreFormat("A. Inic.", 7) & oImpre.ImpreFormat("Not", 3) & oImpre.ImpreFormat("Cal", 3) & oImpre.ImpreFormat("Monto", 10) & oImpre.ImpreFormat("Saldo Cap", 10) & oImpre.ImpreFormat("Moneda", 8) & oImpre.ImpreFormat("Rep", 5) & lnSaltoLinDoc
    'sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 9) & oImpre.ImpreFormat("A.", 5) & oImpre.ImpreFormat("A. Inic.", 7) & oImpre.ImpreFormat("Not", 3) & oImpre.ImpreFormat("Cal", 3) & oImpre.ImpreFormat("Monto", 10) & oImpre.ImpreFormat("Saldo Cap", 10) & oImpre.ImpreFormat("Moneda", 8) & oImpre.ImpreFormat("Rep", 5) & oImpre.ImpreFormat("Seg. Riesgo Credito", 20) & lnSaltoLinDoc'Segmentacion Riesgo por cliente JOEP20201113
    
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    'sImpresion = sImpresion & Space(5) & String(175, "-") & lnSaltoLinDoc 'Segmentacion Riesgo por cliente JOEP20201113
    
    ContLineas = 12
    If pnNumCreditos > 0 Then
        
        For i = 1 To pnNumCreditos
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatCreditos(i, 1), 5)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 2), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 2), 10)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 3), 12)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatCreditos(i, 4)), "AGENCIA", ""), 8)
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 5), 20)
            sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatCreditos(i, 4)), "AGENCIA", ""), 8)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 5), 18)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 6), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 6), 9)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 7), 9)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 8), 7)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 8), 5)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 9), 7)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 9), 7)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 10), 3)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 14), 3)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatCreditos(i, 11)), 8)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatCreditos(i, 11)), 9) 'Monto
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatCreditos(i, 12)), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatCreditos(i, 12)), 9)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(I, 13), 8) & lnSaltoLinDoc
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 13), 8)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 15), 5) & lnSaltoLinDoc
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 15), 5) 'Segmentacion Riesgo por cliente JOEP20201113
            'FIN ORCR-20140913************
            
            'Segmentacion Riesgo por cliente JOEP20201113
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatCreditos(i, 16), 20) & lnSaltoLinDoc
            'Segmentacion Riesgo por cliente JOEP20201113
            
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE CREDITOS : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Tipo", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
                sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Analista", 7) & oImpre.ImpreFormat("A. Inicial", 7) & oImpre.ImpreFormat("Nota", 5) & oImpre.ImpreFormat("Monto", 10) & oImpre.ImpreFormat("Saldo Cap", 10) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 13
            End If
        Next i
        
        'Segmentacion Riesgo por cliente JOEP20201113
        sImpresion = sImpresion & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & "SECCION DE SEGMENTACIÓN RIESGO POR CLIENTE: " & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatCreditos(1, 16), 200) & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        ContLineas = ContLineas + 5
'            If pMatCreditos(1, 16) <> "" Then
'                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatCreditos(1, 17), 25) & lnSaltoLinDoc
'            Else
'                sImpresion = sImpresion & Space(5) & "Cliente No posee segmentación de riesgo" & lnSaltoLinDoc & lnSaltoLinDoc
'                ContLineas = ContLineas + 2
'            End If
        'Segmentacion Riesgo por cliente JOEP20201113
        
    '***LUCV20170302, ANEXO 001-2017 Alertas Tempranas ********************************************************************
        Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        sImpresion = sImpresion & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & "SECCION DE ALERTAS TEMPRANAS : " & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("Nro. Crédito", 25) & oImpre.ImpreFormat("Indicadores", 55) & oImpre.ImpreFormat("Resultado", 20) & oImpre.ImpreFormat("Alerta", 15) & oImpre.ImpreFormat("Valor Limite", 15) & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        ContLineas = ContLineas + 5
            Set R = oDCOMFormatosEval.ObtieneListadoSeccionAlertasTempranas(psPersCod)
            If Not (R.EOF Or R.BOF) Then
                Do While Not R.EOF
                    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!cCtaCod, 18, , False)
                    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!cAlerta, 40, , False)
                    sImpresion = sImpresion & Space(0) & oImpre.ImpreFormat(R!nValor, 20, , False)
                    sImpresion = sImpresion & oImpre.ImpreFormat(R!cUnidad, 10, , False) 'lucv
                    
                    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!cRequiere, 8, , False)
                    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!cComparador, 2, , False) 'lucv
                    sImpresion = sImpresion & oImpre.ImpreFormat(R!nValorLimite, 6, , True)
                    sImpresion = sImpresion & oImpre.ImpreFormat(R!cUnidad, 10, , False) & lnSaltoLinDoc 'lucv
                    ContLineas = ContLineas + 1
                    If ContLineas > 56 Then
                        sImpresion = sImpresion & lnSaltoPagina
                        sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & "SECCION DE ALERTAS TEMPRANAS : " & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("Nro. Crédito", 25) & oImpre.ImpreFormat("Indicadores", 68) & oImpre.ImpreFormat("Resultado", 12) & oImpre.ImpreFormat("Alerta", 12) & oImpre.ImpreFormat("Valor Limite", 13) & lnSaltoLinDoc
                        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                        ContLineas = 13
                    End If
                    R.MoveNext
                Loop
            Else
                sImpresion = sImpresion & Space(5) & "Cliente No posee ningun indicador de alertas tempranas" & lnSaltoLinDoc & lnSaltoLinDoc
                ContLineas = ContLineas + 2
            End If
        Set oCred = Nothing
        Set R = Nothing
    'Fin LUCV20170302
    
        '--------------GARANTIAS----------------------
        'Set R = New ADODB.Recordset
        Set oCred = New COMDCredito.DCOMCredito
        sImpresion = sImpresion & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & "SECCION DE GARANTIAS : " & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("MONTO REALIZ.", 13) & oImpre.ImpreFormat("MONTO DISPON.", 13) & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        ContLineas = ContLineas + 5
            Set R = oCred.RecuperaGarantiasCreditoConsol(psPersCod, pdFecSis)
            Do While Not R.EOF
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!nRealizacion, 8, , True)
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(R!nPorGravar, 8, , True)
                ContLineas = ContLineas + 1
                If ContLineas > 56 Then
                    sImpresion = sImpresion & lnSaltoPagina
                    sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & "SECCION DE GARANTIAS : " & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("MONTO REALIZ.", 13) & oImpre.ImpreFormat("MONTO DISPON.", 13) & lnSaltoLinDoc
                    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                    ContLineas = 13
                End If
                R.MoveNext
            Loop
        Set oCred = Nothing
        Set R = Nothing
        
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Cuentas de Creditos" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
    '********** A H O R R O S **********************
    sImpresion = sImpresion & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE AHORROS : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(160, "-") & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Producto", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
    sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10)
    If pnSaldosAho = 1 Then
        sImpresion = sImpresion & oImpre.ImpreFormat("Saldo Cont", 10) & Space(4) & oImpre.ImpreFormat("Saldo Disp", 10) & oImpre.ImpreFormat("Bloq. Parcial", 12)
    End If
    sImpresion = sImpresion & oImpre.ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(160, "-") & lnSaltoLinDoc
    
    ContLineas = ContLineas + 6
    If pnNumAhorros > 0 Then
        For i = 1 To pnNumAhorros
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatAhorros(i, 1), 5)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 2), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 3), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatAhorros(i, 4)), "AGENCIA", ""), 8)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 5), 20)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 6), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 7), 12)
            If pnSaldosAho = 1 Then
                sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 8), 12, 2)
                sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 9), 12, 2)
                sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 10), 12, 2)
            End If
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 11), 17)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatAhorros(i, 12), 10) & lnSaltoLinDoc
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE AHORROS : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Producto", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
                sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10)
                If pnSaldosAho = 1 Then
                    sImpresion = sImpresion & oImpre.ImpreFormat("SaldoCont", 10) & Space(4) & oImpre.ImpreFormat("Saldo Dispon", 8)
                End If
                sImpresion = sImpresion & oImpre.ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 13
            End If
        Next i
        
        sImpresion = sImpresion & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & "TOTAL AHORROS:   SOLES=" & Trim(psSolesAho) & Space(10 - Len(psSolesAho)) & " DOLARES=" & Trim(psDolaresAho) & lnSaltoLinDoc
        
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Cuentas de Ahorros" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
    'Para Pignoraticio
    sImpresion = sImpresion & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
    sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("No Renov.", 10) & oImpre.ImpreFormat("Monto", 6) & oImpre.ImpreFormat("Fecha Venc", 12) & oImpre.ImpreFormat("Saldo Cap.", 10) & oImpre.ImpreFormat("Tasación", 10) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    ContLineas = ContLineas + 6
    If pnNumPrendario > 0 Then
        For i = 1 To pnNumPrendario
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatPrendario(i, 1), 5)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 2), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatPrendario(i, 3)), "AGENCIA", ""), 8)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 4), 20)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 5), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 6), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 7), 3)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatPrendario(i, 8)), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 9), 11)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatPrendario(i, 10)), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatPrendario(i, 11)), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatPrendario(i, 12), 11) & lnSaltoLinDoc
            
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 10) & oImpre.ImpreFormat("Estado", 10)
                sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("No Renov.", 10) & oImpre.ImpreFormat("Monto", 8) & oImpre.ImpreFormat("Fecha Venc", 10) & oImpre.ImpreFormat("Saldo Cap.", 10) & oImpre.ImpreFormat("Tasación", 10) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 13
            End If
        Next i
        sImpresion = sImpresion & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & "TOTAL PIGNORATICIO:   SOLES=" & Trim(psSolesPig) & Space(10 - Len(psSolesPig)) & " DOLARES=" & Trim(psDolaresPig) & lnSaltoLinDoc
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Cuentas de Pignoraticio" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
    'Judicial
    sImpresion = sImpresion & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE JUDICIAL : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    'INICIO ORCR-20140913*********
    'sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Tipo", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 13)
    'sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Saldo Cap.", 10) & oImpre.ImpreFormat("Fecha Cast.", 10) & oImpre.ImpreFormat("Fecha Ult. Pago", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 10) & oImpre.ImpreFormat("Tipo", 12) & oImpre.ImpreFormat("Agencia", 8) & oImpre.ImpreFormat("Cuenta", 18) & oImpre.ImpreFormat("Estado", 13)
    sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Saldo Cap.", 10) & oImpre.ImpreFormat("Rep", 5) & oImpre.ImpreFormat("Fecha Cast.", 10) & oImpre.ImpreFormat("Fecha Ult. Pago", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    ContLineas = ContLineas + 4
    
    If pnNumJudicial > 0 Then
        For i = 1 To pnNumJudicial
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatJudicial(i, 1), 5)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 2), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 2), 10)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 3), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatJudicial(i, 4)), "AGENCIA", ""), 8)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 5), 20)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 5), 18)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 6), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 6), 13)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 7), 9)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatJudicial(i, 8)), 8) 'Saldo Cap
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatJudicial(i, 8)), 10) 'Saldo Cap
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 9), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 9), 5)
            
            'sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 10), 10) & lnSaltoLinDoc
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 10), 10)
            
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatJudicial(i, 11), 10) & lnSaltoLinDoc
            'FIN ORCR-20140913************
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE JUDICIAL : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Tipo", 10) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 10) & oImpre.ImpreFormat("Estado", 10)
                sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Saldo Cap.", 10) & oImpre.ImpreFormat("Fecha Cast.", 10) & oImpre.ImpreFormat("Fecha Ult. Pago", 10) & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 13
            End If
        Next i
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Cuentas en Judicial" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
    '***************** CARTA FIANZA ************************
    sImpresion = sImpresion & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
    sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Analista", 10) & oImpre.ImpreFormat("Monto", 10) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    ContLineas = ContLineas + 6
    If pnNumCartaFianza > 0 Then
        
        For i = 1 To pnNumCartaFianza
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatCartaFianza(i, 1), 5)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 2), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(Replace(UCase(pMatCartaFianza(i, 3)), "AGENCIA", ""), 8)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 4), 20)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 5), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 6), 12)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 7), 3)
            sImpresion = sImpresion & oImpre.ImpreFormat(CDbl(pMatCartaFianza(i, 8)), 10)
            sImpresion = sImpresion & oImpre.ImpreFormat(pMatCartaFianza(i, 9), 8) & lnSaltoLinDoc
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("NRO", 5) & oImpre.ImpreFormat("Fecha", 12) & oImpre.ImpreFormat("Agencia", 10) & oImpre.ImpreFormat("Cuenta", 20) & oImpre.ImpreFormat("Estado", 10)
                sImpresion = sImpresion & oImpre.ImpreFormat("Particip", 10) & oImpre.ImpreFormat("Analista", 10) & oImpre.ImpreFormat("Monto", 15) & oImpre.ImpreFormat("Moneda", 10) & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 13
            End If
            
        Next i
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Cartas Fianza" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
    '***************** COMENTARIOS ************************
    Dim sCadena As String, sLinea As String
    Dim nPos As Integer
    sImpresion = sImpresion & lnSaltoLinDoc & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE COMENTARIO : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    ContLineas = ContLineas + 5
    sCadena = Trim(psComentario)
    If sCadena <> "" Then
        Do
            nPos = InStr(1, sCadena, lnSaltoLinDoc, vbTextCompare)
            If nPos <> 0 Then
                sLinea = Mid(sCadena, 1, nPos)
                sCadena = Mid(sCadena, nPos + 1, Len(sCadena) - nPos)
            Else
                sLinea = sCadena
                sCadena = ""
            End If
            sImpresion = sImpresion & Space(5) & sLinea
            
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                sImpresion = sImpresion & lnSaltoPagina
                sImpresion = sImpresion & Space(5) & psNomCmac & Space(90) & "Fecha : " & sFechaHora & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & psNomAge & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & lnSaltoLinDoc & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Cliente : " & psPersCod & Space(2) & psNomPers & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "Documentos :   DNI :" & psDocNat & Space(2) & "RUC :" & psDocJur & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & "SECCION DE COMENTARIOS : " & lnSaltoLinDoc
                sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                ContLineas = 11
            End If
                
        Loop Until sCadena = ""
    Else
        sImpresion = sImpresion & Space(5) & "Cliente No Posee Comentarios" & lnSaltoLinDoc & lnSaltoLinDoc
        ContLineas = ContLineas + 2
    End If
    
'ARCV 25-10-2006
    Dim oDCreDoc As COMDCredito.DCOMCredDoc
    Set oDCreDoc = New COMDCredito.DCOMCredDoc
    Dim RsC As ADODB.Recordset
    Set RsC = New ADODB.Recordset
    Set RsC = oDCreDoc.ObtenerCalificacionDiasAtraso_Titular(psPersCod)
    Set oDCreDoc = Nothing
    If Not RsC.EOF And Not RsC.BOF Then
    psCalifiFinal = RsC!cCalificacion
    End If
    RsC.Close
    Set RsC = Nothing
    'Para Sistema Financiero
    sImpresion = sImpresion & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & "SECCION DE SISTEMA FINANCIERO : " & lnSaltoLinDoc
    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    'ALPA* 20080819************************************************************************
    'For i = 0 To UBound(pMatSistFinac) - 1
    For i = 0 To UBound(pMatSistFinac)
        'If i = 1 Or i = 4 Or i = 5 Or i = 7 Then
        For J = 0 To 4
            If i <= UBound(pMatSistFinac) - 1 Then
                If Trim(pMatSistFinac(i, J)) = "Normal" Or Trim(pMatSistFinac(i, J)) = "Endeudamiento" Or Trim(pMatSistFinac(i, J)) = "Calificacion CMAC - Riesgos" Or Trim(pMatSistFinac(i, J)) = "Fecha" Or Trim(pMatSistFinac(i, J)) = "Directa Soles" Then
                    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
                End If
            End If
            If J = 1 And i = 9 Then
            sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(psCalifiFinal, 18)
            Else
            If i <= UBound(pMatSistFinac) - 1 Then
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat(pMatSistFinac(i, J), 18)
            Else
                sImpresion = sImpresion & Space(5) & oImpre.ImpreFormat("", 18)
            End If
            End If
        Next J
        sImpresion = sImpresion & lnSaltoLinDoc
        'If i = 7 Then
        '    sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
        'End If
    Next i
    '*************************************************************************************
    '***JGPA20181011 ACTA 034-2018
        sImpresion = sImpresion & Space(50) & Chr$(27) + Chr$(71) & "****** ESTE DOCUMENTO NO REPRESENTA UN HISTORIAL DE PAGOS; ES UN DOCUMENTO INTERNO ******" & Chr$(27) + Chr$(72) & lnSaltoLinDoc
        sImpresion = sImpresion & Space(5) & String(150, "-") & lnSaltoLinDoc
    '***JGPA20181011
'----------------
    GeneraImpresionPosicionCliente = sImpresion
    Exit Function
    
ErrorGeneraImpresionPosicionCliente:
    Err.Raise Err.Number, "Genera Impresion Posicion Cliente", Err.Description
End Function
Public Function ImpreRepor_HistorialCredito(ByVal psCtaCod As String, ByVal psNomAge As String, _
                                            ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim i As Integer
    Dim sCadImp As String
    Dim nTotal As Integer
    Dim rs As ADODB.Recordset
    Dim rTasCosEfeAnu As ADODB.Recordset
    Dim oDCredDoc As COMDCredito.DCOMCredDoc
    Dim nMontoCapital As Double
    Dim nMontoInteres As Double
    Dim nMontoMora As Double
    Dim nMontoGastos As Double
    Dim nPagoTotal As Double '*****YIHUERS1692014
    Dim nPagoTotalDet As Double '*****YIHUERS1692014
    Dim nTasaEfectivaAnual, nTasaCostoEfectivoAnual As Double
    'Dim rs As ADODB.Recordset
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "HISTORIAL DEL CREDITO", psNomCmac)
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set rs = oDCredDoc.Repor_HistorialCreditoCabecera(psCtaCod)
    Set oDCredDoc = Nothing
    
    Dim oCred As COMNCredito.NCOMCredito
    Set oCred = New COMNCredito.NCOMCredito
    
    If Not rs Is Nothing Then
        If Not (rs.EOF And rs.BOF) Then
            nTasaEfectivaAnual = oCred.TasaIntPerDias(rs!TasaInteres, 360) * 100
        End If
    End If
    Set oCred = Nothing
    
    
    Dim oDCred As COMDCredito.DCOMCredActBD
    Set oDCred = New COMDCredito.DCOMCredActBD
    
    Set rTasCosEfeAnu = oDCred.CargaRecordSet("select nTasCosEfeAnu from ColocacCred where cCtaCod='" & psCtaCod & "'")
    
    If Not rTasCosEfeAnu Is Nothing Then
        If rTasCosEfeAnu.EOF And rTasCosEfeAnu.BOF Then
            nTasaCostoEfectivoAnual = 0
        Else
            nTasaCostoEfectivoAnual = IIf(IsNull(rTasCosEfeAnu!nTasCosEfeAnu), 0, rTasCosEfeAnu!nTasCosEfeAnu)
        End If
    Else
        nTasaCostoEfectivoAnual = 0
    End If
        
    If Not rs.EOF And Not rs.BOF Then
        sCadImp = sCadImp & oImpre.ImpreFormat("NRO CREDITO: ", 10) & oImpre.ImpreFormat(psCtaCod, 20) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat("CLIENTE: ", 10) & oImpre.ImpreFormat(rs!cPersNombre, 20) & oImpre.ImpreFormat("CODIGO: ", 15) & oImpre.ImpreFormat(rs!cPersCod, 15) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat("CREDITO: ", 10) & oImpre.ImpreFormat(psCtaCod, 20) & oImpre.ImpreFormat("VIGENCIA: ", 15) & oImpre.ImpreFormat(Format(rs!dVigencia, "dd/MM/yyyy"), 15) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat("ESTADO: ", 10) & oImpre.ImpreFormat(rs!cEstado, 20) & oImpre.ImpreFormat("VENCIMIENTO: ", 15) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 15) & lnSaltoLinDoc
        'sCadImp = sCadImp & oImpre.ImpreFormat("TASA: ", 10) & oImpre.ImpreFormat(rs!TasaInteres & "%", 20, 2) & oImpre.ImpreFormat("ANALISTA: ", 15) & oImpre.ImpreFormat(rs!cUser, 15) & lnSaltoLinDoc
        'sCadImp = sCadImp & oimpre.ImpreFormat("DESEMB: ", 10) & oimpre.ImpreFormat(rs!nMontoCol, 20, 2) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat("TEA: ", 10) & oImpre.ImpreFormat(Format(nTasaEfectivaAnual, "#0.00") & "%", 20, 2) & oImpre.ImpreFormat("TCEA: ", 15) & oImpre.ImpreFormat(Format(nTasaCostoEfectivoAnual, "#0.00") & "%", 15) & lnSaltoLinDoc
        sCadImp = sCadImp & oImpre.ImpreFormat("IMPORTE: ", 10) & Space(2) & Format(rs!nMontoCol, "#0.00") & lnSaltoLinDoc
    End If
    Set rs = Nothing
    
    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set rs = oDCredDoc.Repo_HistorialCreditoDetalle(psCtaCod)
    Set oDCredDoc = Nothing
    
    'imprimiendo las columnas del detalle
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & String(145, "-") & Chr(10) '*****YIHUERS1692014
    'sCadImp = sCadImp & oImpre.ImpreFormat("NROCUOTA", 10) & oImpre.ImpreFormat("FEC.VENC", 10) & oImpre.ImpreFormat("FEC.PAGO", 10) & oImpre.ImpreFormat("CAPITAL", 8) & oImpre.ImpreFormat("INTERES", 8) & Space(6) & oImpre.ImpreFormat("MORA", 8) & oImpre.ImpreFormat("GASTOS", 8) & oImpre.ImpreFormat("Int. Gracia", 13) & oImpre.ImpreFormat("ATRASO", 8) & oImpre.ImpreFormat("Saldo Capital", 13) & oImpre.ImpreFormat("USUARIO", 8) & oImpre.ImpreFormat("PAGO TOTAL", 10) & lnSaltoLinDoc '*****YIHUERS1692014
    sCadImp = sCadImp & oImpre.ImpreFormat("NROCUOTA", 10) & oImpre.ImpreFormat("FEC.VENC", 10) & oImpre.ImpreFormat("FEC.PAGO", 10) & oImpre.ImpreFormat("CAPITAL", 8) & oImpre.ImpreFormat("INTERES", 8) & Space(6) & oImpre.ImpreFormat("MORA", 8) & oImpre.ImpreFormat("GASTOS", 8) & oImpre.ImpreFormat("Int. Gracia", 13) & oImpre.ImpreFormat("ATRASO", 8) & oImpre.ImpreFormat("PAGO TOTAL", 13) & oImpre.ImpreFormat("Saldo Capital", 13) & oImpre.ImpreFormat("USUARIO", 8) & lnSaltoLinDoc '*****RECO20151111 ERS039-2015
    sCadImp = sCadImp & String(145, "-") & Chr(10) '*****YIHUERS1692014
    
    nMontoCapital = 0
    nMontoInteres = 0
    nMontoMora = 0
    nMontoGastos = 0
    nPagoTotal = 0 '*****YIHUERS1692014
    Do Until rs.EOF
        i = i + 1
        nPagoTotalDet = rs!Interes + rs!Mora + rs!Gastos + rs!IntGra + IIf(rs!nColocCalendApl = 1, rs!Capital, 0) '*****YIHUERS1692014
        nPagoTotal = nPagoTotal + nPagoTotalDet '*****YIHUERS1692014
        
        'sCadImp = sCadImp & oImpre.ImpreFormat(rs!nNroCuota, 10, 0) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!Fechapago, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(rs!Capital, 8, 2) & oImpre.ImpreFormat(rs!Interes, 8, 2) & oImpre.ImpreFormat(rs!Mora, 8, 2) & oImpre.ImpreFormat(rs!Gastos, 8, 2) & oImpre.ImpreFormat(rs!IntGra, 13, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8) & oImpre.ImpreFormat(rs!nSaldoCap, 14, 2) & oImpre.ImpreFormat(rs!cUsuario, 8) & lnSaltoLinDoc '*****YIHUERS1692014
        'sCadImp = sCadImp & oImpre.ImpreFormat(rs!nNroCuota, 10, 0) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!Fechapago, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(rs!Capital, 8, 2) & oImpre.ImpreFormat(rs!Interes, 8, 2) & oImpre.ImpreFormat(rs!Mora, 8, 2) & oImpre.ImpreFormat(rs!Gastos, 8, 2) & oImpre.ImpreFormat(rs!IntGra, 13, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8) & oImpre.ImpreFormat(rs!nSaldoCap, 14, 2) & oImpre.ImpreFormat(rs!cUsuario, 8) & oImpre.ImpreFormat(nPagoTotalDet, 10) & lnSaltoLinDoc '*****YIHUERS1692014
        sCadImp = sCadImp & oImpre.ImpreFormat(rs!nNroCuota, 10, 0) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(rs!Fechapago, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(rs!Capital, 8, 2) & oImpre.ImpreFormat(rs!Interes, 8, 2) & oImpre.ImpreFormat(rs!Mora, 8, 2) & oImpre.ImpreFormat(rs!Gastos, 8, 2) & oImpre.ImpreFormat(rs!IntGra, 13, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8, 0) & oImpre.ImpreFormat(nPagoTotalDet, 14, 2) & oImpre.ImpreFormat(rs!nSaldoCap, 14, 2) & Space(1) & oImpre.ImpreFormat(rs!cUsuario, 8) & lnSaltoLinDoc '*****RECO20151111 ERS039-2015
        If i > 1 Then
            nMontoCapital = nMontoCapital + rs!Capital
            nMontoInteres = nMontoInteres + rs!Interes
            nMontoMora = nMontoMora + rs!Mora
            nMontoGastos = nMontoGastos + rs!Gastos
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set rs = oDCredDoc.Repo_HistorialCredDetalleComplem(psCtaCod, pdFecSis)
    Set oDCredDoc = Nothing
    
    Do Until rs.EOF
        nPagoTotalDet = rs!Interes + rs!Mora + rs!Gastos + rs!IntGra '*****YIHUERS1692014
        nPagoTotal = nPagoTotal + nPagoTotalDet '*****YIHUERS1692014
        'sCadImp = sCadImp & oImpre.ImpreFormat(rs!nCuota, 10, 0) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(IIf(IsNull(rs!Fechapago), "", rs!Fechapago), "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(rs!Captital, 8, 2) & oImpre.ImpreFormat(rs!Interes, 8, 2) & oImpre.ImpreFormat(rs!Mora, 8, 2) & oImpre.ImpreFormat(rs!Gastos, 8, 2) & oImpre.ImpreFormat(rs!IntGra, 13, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8) & oImpre.ImpreFormat(rs!nSaldoCap, 14, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!cUsuario), "", rs!cUsuario), 8) & lnSaltoLinDoc '*****YIHUERS1692014
        sCadImp = sCadImp & oImpre.ImpreFormat(rs!nCuota, 10, 0) & oImpre.ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(Format(IIf(IsNull(rs!Fechapago), "", rs!Fechapago), "dd/MM/yyyy"), 10) & oImpre.ImpreFormat(rs!Captital, 8, 2) & oImpre.ImpreFormat(rs!Interes, 8, 2) & oImpre.ImpreFormat(rs!Mora, 8, 2) & oImpre.ImpreFormat(rs!Gastos, 8, 2) & oImpre.ImpreFormat(rs!IntGra, 13, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8, 0) & oImpre.ImpreFormat(rs!nSaldoCap, 14, 2) & oImpre.ImpreFormat(IIf(IsNull(rs!cUsuario), "", rs!cUsuario), 8) & oImpre.ImpreFormat(nPagoTotalDet, 8) & lnSaltoLinDoc '*****YIHUERS1692014
        nPagoTotalDet = 0
        rs.MoveNext
    Loop
    
    'sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc '*****YIHUERS1692014
    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc '*****YIHUERS1692014
    'sCadImp = sCadImp & Space(35) & oImpre.ImpreFormat(nMontoCapital, 8, 2) & oImpre.ImpreFormat(nMontoInteres, 8, 2) & oImpre.ImpreFormat(nMontoMora, 8, 2) & oImpre.ImpreFormat(nMontoGastos, 8, 2) & lnSaltoLinDoc '*****YIHUERS1692014
    sCadImp = sCadImp & Space(35) & oImpre.ImpreFormat(nMontoCapital, 8, 2) & oImpre.ImpreFormat(nMontoInteres, 8, 2) & oImpre.ImpreFormat(nMontoMora, 8, 2) & oImpre.ImpreFormat(nMontoGastos, 8, 2) & Space(29) & oImpre.ImpreFormat(nPagoTotal, 8, 2) & lnSaltoLinDoc '*****YIHUERS1692014
    'sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc '*****YIHUERS1692014
    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc '*****YIHUERS1692014
    ImpreRepor_HistorialCredito = sCadImp
End Function

Public Function ImprimeConsultaCredito(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
                                ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psEstado As String, ByVal psLineaCred As String, ByVal psFteIngreso As String, _
                                ByVal MatRelacCred As Variant, ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, _
                                ByVal psApoderado As String, ByVal psTipoCuota As String, ByVal MatHitorial As Variant, ByVal MatDesembolsos As Variant, _
                                ByVal MatPagos As Variant, ByVal MatCuotasPend As Variant, ByVal MatDeudaVenc As Variant, ByVal MatDeudaAFecha As Variant, ByVal psProtesto As String, _
                                ByVal psCargoAutom As String, ByVal psRefinanciado As String, ByVal psNomCmac As String, ByVal pMatGarantias As Variant, _
                                ByVal pnCuotaCom As Integer, ByVal pnMiViv As Integer, ByVal pnCalendDin As Integer, ByVal psTasaComp As String, ByVal psTasaMor As String, ByVal MatRefinan As Variant, ByVal psFecEEFF As String, _
                                Optional ByVal pnTea As String, Optional ByVal pnTcea As String, Optional ByVal mFinanciado As String = "", Optional ByVal psMensajeMIVIVIENDA As String = "") As String
                                'peac 20071228 se agrego "ByVal psFecEEFF As String"
                                ' RIRO20131202 Se agregó "pnTea, pnTcea" Segun ERS098-2013
                                'WIOR 20151224 AGREGO psMensajeMIVIVIENDA
     
Dim sCadImp As String
Dim i As Integer

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Resumen General de Credito", psNomCmac, Space(5))
    sCadImp = sCadImp & lnSaltoLinDoc & Space(5) & lsNegritaOn & "Credito : " & psCtaCod & lsNegritaOff & Space(20) & "Tipo : " & psTipoCred & Space(20) & "Estado : " & psEstado & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CLIENTES RELACIONADOS" & Space(45) & "DATOS GENERALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(5) & String(60, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "Nombre" & Space(35) & "Relaciones" & Space(15) & "Linea de Credito      : " & psLineaCred & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(6) & "Fuente de Ingreso     : " & psFteIngreso & lnSaltoLinDoc
    
    'RIRO20131202 ERS098-2013 Se aumento bucle a 7
    For i = 0 To 7
        sCadImp = sCadImp & Space(5)
        If i <= UBound(MatRelacCred) - 1 Then
            sCadImp = sCadImp & oImpre.ImpreFormat(MatRelacCred(i, 0), 40, 0) & oImpre.ImpreFormat(MatRelacCred(i, 1), 20, 0)
        End If
        Select Case i
            Case 0
                sCadImp = sCadImp & Space(6) & "Analista Responsable  : " & psAnalista & lnSaltoLinDoc
            Case 1
                sCadImp = sCadImp & Space(6) & "Condicion del Credito : " & psCondicion & lnSaltoLinDoc
            Case 2
                sCadImp = sCadImp & Space(66) & "Destino del Credito   : " & psDestino & lnSaltoLinDoc
            Case 3
                sCadImp = sCadImp & Space(66) & "Tipo de Cuota         : " & psTipoCuota & lnSaltoLinDoc
            Case 4
                'sCadImp = sCadImp & Space(66) & "Tasa Compensatoria    : " & psTasaComp & lnSaltoLinDoc RIRO20131202
                sCadImp = sCadImp & Space(66) & "Tasa Efectiva Anual   : " & pnTea & lnSaltoLinDoc 'RIRO20131202
            Case 5
                sCadImp = sCadImp & Space(66) & "Tasa Costo Efec. Anual: " & pnTcea & lnSaltoLinDoc 'RIRO20131202
            Case 6
                sCadImp = sCadImp & Space(66) & "Tasa Moratoria        : " & psTasaMor & lnSaltoLinDoc
            Case 7
                'peac 20071228
                sCadImp = sCadImp & Space(66) & "Ultima Fecha de EE.FF : " & psFecEEFF & lnSaltoLinDoc
        End Select
    Next i
    
    'Hitorial de Credito
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "HISTORIAL DE CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & Space(10) & oImpre.ImpreFormat("FECHA", 12) & oImpre.ImpreFormat("MONTO", 12) & oImpre.ImpreFormat("Nro CUOTAS", 12) & oImpre.ImpreFormat("PLAZO", 12) & oImpre.ImpreFormat("CUOTA", 12) & oImpre.ImpreFormat("P. GRACIA", 12) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For i = 0 To UBound(MatHitorial) - 1
        '                   Estado                                  Fecha                                   Monto                                   Nro Cuotas                              Plazo                                   Cuota                                   Periodo de Gracia
        sCadImp = sCadImp & Space(5) & oImpre.ImpreFormat(MatHitorial(i, 0), 11, 0) & oImpre.ImpreFormat(MatHitorial(i, 1), 16, 0) & oImpre.ImpreFormat(MatHitorial(i, 2), 12, 0) & oImpre.ImpreFormat(MatHitorial(i, 3), 16, 0) & oImpre.ImpreFormat(MatHitorial(i, 4), 12, 0) & oImpre.ImpreFormat(MatHitorial(i, 5), 18, 0) & oImpre.ImpreFormat(MatHitorial(i, 6), 12, 0) & lnSaltoLinDoc
    Next i
    
    'DESEMBOLSOS DEL CREDITO
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "DESEMBOLSOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & oImpre.ImpreFormat("FECHA", 18, 0) & oImpre.ImpreFormat("MONTO", 15, 0) & oImpre.ImpreFormat("GASTOS", 12, 0) & oImpre.ImpreFormat("ESTADO", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For i = 0 To UBound(MatDesembolsos) - 1
        '                               FECHA                                       MONTO                                     GASTOS                                    ESTADO DEL DESEMBOLSO
        sCadImp = sCadImp & Space(5) & oImpre.ImpreFormat(MatDesembolsos(i, 0), 19, 0) & oImpre.ImpreFormat(MatDesembolsos(i, 1), 14, 0) & oImpre.ImpreFormat(MatDesembolsos(i, 2), 12, 0) & oImpre.ImpreFormat(MatDesembolsos(i, 3), 18, 0) & lnSaltoLinDoc
    Next i
    
    
    
    'CUOTAS PAGADAS
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PAGADAS" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(147, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & oImpre.ImpreFormat("CUOTA", 10, 0) & oImpre.ImpreFormat("FECHA PAG.", 17, 0) & oImpre.ImpreFormat("TOTAL PAG.", 14, 0) & oImpre.ImpreFormat("CAPITAL", 17, 0) & oImpre.ImpreFormat("INTERES", 12, 0) & oImpre.ImpreFormat("MORA", 11, 2) & oImpre.ImpreFormat("GASTOS", 12, 0) & oImpre.ImpreFormat("Int. Gracia", 15, 0) & oImpre.ImpreFormat("ITF", 10, 0) & oImpre.ImpreFormat("ATRASO", 13, 0) & oImpre.ImpreFormat("SALDO.K.", 10, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(147, "-") & lnSaltoLinDoc
    For i = 0 To UBound(MatPagos) - 1
        '                               Nro Cuota                           FECHA DE VENCIMIENTO                       TOTAL PAGADO                              CAPITAL                                   INTERES                                   MORA                                      GASTOS                                    ITF                                       DIAS ATRASO                               Saldo Kapìtal
        sCadImp = sCadImp & Space(10) & oImpre.ImpreFormat(MatPagos(i, 1), 8, 0) & oImpre.ImpreFormat(MatPagos(i, 0), 18, 0) & oImpre.ImpreFormat(MatPagos(i, 2), 15, 0) & oImpre.ImpreFormat(MatPagos(i, 3), 16, 0) & oImpre.ImpreFormat(MatPagos(i, 4), 14, 0) & oImpre.ImpreFormat(MatPagos(i, 5), 9, 0) & oImpre.ImpreFormat(MatPagos(i, 6), 12, 2) & oImpre.ImpreFormat(MatPagos(i, 10), 13, 0) & oImpre.ImpreFormat(MatPagos(i, 7), 10, 2) & oImpre.ImpreFormat(MatPagos(i, 8), 8, 2) & oImpre.ImpreFormat(MatPagos(i, 9), 15, 2) & lnSaltoLinDoc
    Next i
    'ORCR INICIO 20140414 ***
    'MONTO FINANCIADO
    Dim sleasing As String
    sleasing = Mid(psCtaCod, 6, 3)
    If (sleasing = "515" Or sleasing = "516") Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & lsNegritaOn & "MONTO FINANCIADO         : " & mFinanciado & lsNegritaOff & lnSaltoLinDoc
    End If
    'ORCR FIN 20140414 ******
    'CUOTAS PENDIENTES
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PENDIENTES" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & oImpre.ImpreFormat("CUOTA", 8, 0) & oImpre.ImpreFormat("FECHA", 8, 0) & oImpre.ImpreFormat("CAPITAL", 10, 0) & oImpre.ImpreFormat("INTERES", 14, 0) & oImpre.ImpreFormat("MORA", 8, 0) & oImpre.ImpreFormat("DIAS ATRASO", 15, 0) & oImpre.ImpreFormat("ITF", 14, 0) & oImpre.ImpreFormat("CUOTA. ", 14, 0) & oImpre.ImpreFormat("SALDO K.", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For i = 0 To UBound(MatCuotasPend) - 1
        '                                                   Nro Cuota                                FECHA DE VENCIMIENTO                      CAPITAL                                   INTERES                                   MORA                                      DIAS ATRASO                                          ITF                                        SALDO K
        sCadImp = sCadImp & Space(10) & oImpre.ImpreFormat(MatCuotasPend(i, 0), 4, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 1), 12, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 2), 12, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 3), 12, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 4), 12, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 5), 8, 0) & oImpre.ImpreFormat(MatCuotasPend(i, 6), 12, 2) & oImpre.ImpreFormat(MatCuotasPend(i, 8), 12, 2) & oImpre.ImpreFormat(MatCuotasPend(i, 7), 10, 2) & lnSaltoLinDoc
    Next i
    
    'WIOR 20151224 ***
    If Trim(psMensajeMIVIVIENDA) <> "" Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & "Nota:" & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & psMensajeMIVIVIENDA & lnSaltoLinDoc
    End If
    'WIOR FIN ********
    
    'CUOTAS ATRASADAS Y PAGOS
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS ATRASADAS" & Space(45) & "DEUDA A LA FECHA" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "CAPITAL      : " & oImpre.ImpreFormat(CDbl(MatDeudaVenc(0)), 20) & Space(27) & "CAPITAL      : " & oImpre.ImpreFormat(CDbl(MatDeudaAFecha(0)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "INTERES      : " & oImpre.ImpreFormat(CDbl(MatDeudaVenc(1)), 20) & Space(27) & "INTERES      : " & oImpre.ImpreFormat(CDbl(MatDeudaAFecha(1)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "MORA         : " & oImpre.ImpreFormat(CDbl(MatDeudaVenc(2)), 20) & Space(27) & "MORA         : " & oImpre.ImpreFormat(CDbl(MatDeudaAFecha(2)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "GASTOS       : " & oImpre.ImpreFormat(CDbl(MatDeudaVenc(3)), 20) & Space(27) & "GASTOS       : " & oImpre.ImpreFormat(CDbl(MatDeudaAFecha(3)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "TOTAL        : " & oImpre.ImpreFormat(CDbl(MatDeudaVenc(4)), 20) & Space(27) & "TOTAL        : " & oImpre.ImpreFormat(CDbl(MatDeudaAFecha(4)), 20) & lnSaltoLinDoc
    
    'JUEZ 20140414 ******************************************************
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "Nota:" & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "Para el caso de Personas Jurídicas (Medianas, Grandes Empresas así como los créditos Corporativos) se deberá considerar" & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "adicionalmente la comisión por cancelación anticipada según el tarifario vigente." & lnSaltoLinDoc
    'END JUEZ ***********************************************************
    
    'sCadImp = sCadImp & Chr$(12)
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, pscodusu, "Resumen General de Credito", Space(5))
    sCadImp = sCadImp & Space(5) & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "OTROS DATOS" & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(70, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "PROTESTO           : " & psProtesto & Space(10) & oImpre.ImpreFormat("CUOTA COMODIN    : " & IIf(pnCuotaCom = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "CARGO AUTOMATICO   : " & psCargoAutom & Space(10) & oImpre.ImpreFormat("MI VIVIENDA      : " & IIf(pnMiViv = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "REFINANCIADO       : " & psRefinanciado & Space(10) & oImpre.ImpreFormat("CALEND. DIN.     : " & IIf(pnCalendDin = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    'sCadImp = sCadImp & Chr$(12)
    
    'Imprime garantias
    If IsArray(pMatGarantias) Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Resumen General de Credito", psNomCmac, Space(5))
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & " GARANTIAS " & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & Space(10) & oImpre.ImpreFormat("Tipo Garantía", 20) & oImpre.ImpreFormat("Descripción", 20) & oImpre.ImpreFormat("Documento", 15) & oImpre.ImpreFormat("Nro Doc.", 12) & oImpre.ImpreFormat("Moneda", 12) & oImpre.ImpreFormat("Monto", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
        For i = 0 To UBound(pMatGarantias) - 1
            sCadImp = sCadImp & Space(15) & oImpre.ImpreFormat(pMatGarantias(i, 0), 20) & oImpre.ImpreFormat(pMatGarantias(i, 1), 20) & oImpre.ImpreFormat(pMatGarantias(i, 2), 15) & oImpre.ImpreFormat(pMatGarantias(i, 3), 12) & oImpre.ImpreFormat(pMatGarantias(i, 4), 12) & oImpre.ImpreFormat(pMatGarantias(i, 5), 12) & lnSaltoLinDoc
        Next i
    End If
    
    If IsArray(MatRefinan) Then
        If UBound(MatRefinan) > 0 Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & " REFINANCIACIONES " & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & Space(10) & oImpre.ImpreFormat("Credito", 20) & oImpre.ImpreFormat("Capital", 20) & oImpre.ImpreFormat("Interes", 15) & oImpre.ImpreFormat("Otros", 12) & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
            For i = 0 To UBound(MatRefinan) - 1
                sCadImp = sCadImp & Space(15) & oImpre.ImpreFormat(MatRefinan(i, 0), 20) & oImpre.ImpreFormat(MatRefinan(i, 1), 20) & oImpre.ImpreFormat(MatRefinan(i, 2), 15) & oImpre.ImpreFormat(MatRefinan(i, 3), 12) & lnSaltoLinDoc
            Next i
            
        End If
    End If
    
    '***JGPA20181011 ACTA 034-2018
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(50) & lsNegritaOn & "*** ESTE DOCUMENTO NO REPRESENTA UN CRONOGRAMA DE PAGOS OFICIAL; ES UN DOCUMENTO INTERNO ***" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(150, "-") & lnSaltoLinDoc
    '***JGPA20181011
    
    ImprimeConsultaCredito = sCadImp
End Function

Function ImpreKardex(ByVal prs As ADODB.Recordset, ByVal psNomAge, ByVal psFecSis As Date, ByVal psCodUser As String, _
                    ByVal psNomCmac As String, ByVal psCtaCod As String, ByVal psNomCliente As String) As String
    
    Dim sCadImp As String
    Dim i As Integer
    Dim nTotal As Integer
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUser, "KARDEX DEL CREDITO", psNomCmac)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & " AGENCIA: " & psNomAge & Chr(10)
    sCadImp = sCadImp & " CLIENTE: " & psNomCliente & Chr(10)
    sCadImp = sCadImp & " NRO DE CREDITO: " & psCtaCod & Chr(10)
    sCadImp = sCadImp & String(95, "-") & Chr(10)
    sCadImp = sCadImp & Chr(10)
    
    If prs.RecordCount > 0 Then
        nTotal = prs.RecordCount
    '    RaiseEvent IniciaBarra(nTotal)
    Else
        sCadImp = ""
        Exit Function
    End If
    
     sCadImp = sCadImp & oImpre.ImpreFormat("CUOTA", 11) & oImpre.ImpreFormat("MONTO CUOTA", 10) & oImpre.ImpreFormat("CAPITAL", 10) & oImpre.ImpreFormat("INTERES", 15) & oImpre.ImpreFormat("MORA", 10) & oImpre.ImpreFormat("GASTOS", 10) & oImpre.ImpreFormat("ATRASO", 10, 0) & Chr(10)
     sCadImp = sCadImp & String(95, "-")
     
     sCadImp = sCadImp & Chr(10)
    Do Until prs.EOF
        i = i + 1
        sCadImp = sCadImp & oImpre.ImpreFormat(prs!nCuota, 6) & oImpre.ImpreFormat(prs!MontoCuota, 10, 2) & oImpre.ImpreFormat(prs!Capital, 10, 2) & oImpre.ImpreFormat(prs!Interes, 10, 2) & oImpre.ImpreFormat(prs!Mora, 10) & oImpre.ImpreFormat(prs!Gastos, 10) & oImpre.ImpreFormat(prs!Atraso, 10) & lnSaltoLinDoc
    '    RaiseEvent ProgresoBarra(i, " Kardex de Credito", "Loading...")
        prs.MoveNext
    Loop
    'RaiseEvent FinalizaBarra
    ImpreKardex = sCadImp
End Function

Public Function ImprimeKardexConsultaCredito(ByVal psCtaCod As String, _
                                            ByVal psNomAge As String, _
                                            ByVal pdFecSis As Date, _
                                            ByVal psCodUser As String, _
                                            ByVal psNomCmac As String) As String
    
    Dim objDCRedito As COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    Dim sCadImp As String
    Dim objDVisualizacion As COMNCredito.NCOMVisualizacion
    Dim sTitular As String
    
    Set objDCRedito = New COMDCredito.DCOMCredito
    Set rs = objDCRedito.ListaKardex(psCtaCod)
    Set objDCRedito = Nothing
    
    Set objDVisualizacion = New COMNCredito.NCOMVisualizacion
    sTitular = objDVisualizacion.ObtenerTitularByCredito(psCtaCod)
    Set objDVisualizacion = Nothing
    
    sCadImp = ImpreKardex(rs, psNomAge, pdFecSis, psCodUser, psNomCmac, psCtaCod, sTitular)
    
    ImprimeKardexConsultaCredito = sCadImp
    
End Function

Private Function lnSaltoPagina() As String
    lnSaltoPagina = oImpresora.gPrnSaltoPagina
End Function

Public Function GeneraReporteCargoAutomatico(ByVal MatRep As Variant, ByVal psNomAge As String, ByVal pdFecSis As Date, _
                                            ByVal psNomCmac As String, ByVal pbMonSoles As Boolean, ByVal pnTipCambioV As Double, _
                                            ByVal pnTipCambioC As Double, ByVal pnNumRepo As Integer) As String
Dim sCadImp As String
Dim i As Integer
Dim nSuma1 As Double
Dim nSuma2 As Double
Dim nSuma3 As Double

    sCadImp = Chr$(10)
    sCadImp = sCadImp & Space(5) & psNomAge & Chr$(10)
    sCadImp = sCadImp & Space(5) & Format(pdFecSis, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm:ss") & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & Space(30) & UCase(psNomCmac) & Chr$(10)
    sCadImp = sCadImp & oImpre.PrnSet("B+") & Space(20) & "REPORTE DE AMORTIZACIONES POR CARGO AUTOMATICO" & Chr$(10) & oImpre.PrnSet("B-") & Chr$(10)
    sCadImp = sCadImp & Space(16) & "AMORTIZACION DE CREDITOS EN " & pbMonSoles & " AL " & pdFecSis & Chr$(10) & Chr$(10)
    If pbMonSoles = False Then
        sCadImp = sCadImp & Space(20) & "TIPO DE CAMBIO : VENTA =" & Format(pnTipCambioV, "#0.0000") & " COMPRA=" & Format(pnTipCambioC, "#0.0000") & Chr$(10)
    End If
    sCadImp = sCadImp & String(88, "-") & Chr$(10)
    sCadImp = sCadImp & oImpre.ImpreFormat("CREDITO", 20) & oImpre.ImpreFormat("TITULAR", 33) & oImpre.ImpreFormat("MONTO", 8) & oImpre.ImpreFormat("PAGO", 8) & oImpre.ImpreFormat("SALDO", 8) & Chr$(10)
    sCadImp = sCadImp & String(88, "-") & Chr$(10)
    nSuma1 = 0
    nSuma2 = 0
    For i = 0 To pnNumRepo - 1
        sCadImp = sCadImp & oImpre.ImpreFormat(MatRep(i, 0), 20) & oImpre.ImpreFormat(MatRep(i, 1), 30) & oImpre.ImpreFormat(CDbl(MatRep(i, 2)), 8) & oImpre.PrnSet("B+") & oImpre.ImpreFormat(CDbl(MatRep(i, 3)), 8) & oImpre.PrnSet("B-") & oImpre.ImpreFormat(CDbl(MatRep(i, 4)), 6) & Chr$(10)
        nSuma1 = nSuma1 + CDbl(MatRep(i, 2))
        nSuma2 = nSuma2 + CDbl(MatRep(i, 3))
        nSuma3 = nSuma3 + CDbl(MatRep(i, 4))
    Next i
    sCadImp = sCadImp & String(88, "-") & Chr$(10)
    sCadImp = sCadImp & oImpre.ImpreFormat(nSuma1, 62) & oImpre.ImpreFormat(nSuma2, 8) & oImpre.ImpreFormat(nSuma3, 6) & Chr$(10)
    sCadImp = sCadImp & String(88, "-") & Chr$(10)

    GeneraReporteCargoAutomatico = sCadImp

End Function

Public Function ImprimirMetasAnalistas(ByVal psNomAge As String, _
                                        ByVal pdFechaHora As Date, ByVal psCodUsu As String, ByVal psNomCmac As String)

Dim sCadImp As String
Dim RAnalista As ADODB.Recordset
Dim RDatos As ADODB.Recordset
Dim oCredGen As COMDCredito.DCOMCredito
Dim oDCred As COMDCredito.DCOMCredito
Dim i As Integer
Dim oBase As COMDCredito.DCOMCredActBD
Dim dFecSis As Date
Dim sCadProd As String

    On Error GoTo ErrorImprimirMetasAnalistas
    Set oBase = New COMDCredito.DCOMCredActBD
    dFecSis = oBase.dFechaHora
    Set oBase = Nothing
    dFecSis = CDate(Format(Format(pdFechaHora, "dd/mm/yyyy") & " " & Format(dFecSis, "hh:mm:ss"), "dd/mm/yyyy hh.mm:ss"))
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, Format(dFecSis, "dd/mm/yyyy hh:mm:ss"), psCodUsu, "LISTADO DE METAS POR ANALISTA", psNomCmac)
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & oImpre.ImpreFormat("Tipo", 12, 0) & oImpre.ImpreFormat("Fec Inicio", 14, 0) & oImpre.ImpreFormat("Fec Final", 15, 0) & Space(9)
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set RDatos = oDCred.RecuperaProductosDeCredito
    Set oDCred = Nothing
    Do While Not RDatos.EOF
        sCadImp = sCadImp & oImpre.ImpreFormat(Trim(RDatos!cAbrev), 10, 1)
        RDatos.MoveNext
    Loop
    RDatos.Close
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    
    Set oCredGen = New COMDCredito.DCOMCredito
    Set RAnalista = oCredGen.CargaAnalistas
    Set oCredGen = Nothing
    Do While Not RAnalista.EOF
        Set oDCred = New COMDCredito.DCOMCredito
        Set RDatos = oDCred.RecuperaDatosMetasAnalistaReporte(RAnalista!cPersCod)
        Set oDCred = Nothing
        
        If Not RDatos.EOF And Not RDatos.BOF Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsNegritaOn
            sCadImp = sCadImp & "ANALISTA : " & oCad.PstaNombre(RAnalista!cPersNombre)
            sCadImp = sCadImp & lsNegritaOff & lnSaltoLinDoc
        End If
        
        Do While Not RDatos.EOF
            sCadImp = sCadImp & oImpre.ImpreFormat(Trim(RDatos!cTipoMeta), 7) & oImpre.ImpreFormat(Format(RDatos!dInicial, "dd/mm/yyyy"), 12) _
                & oImpre.ImpreFormat(Format(RDatos!dfinal, "dd/mm/yyyy"), 12) & oImpre.ImpreFormat(Trim(RDatos!sMoneda), 8)
            'Imprimir los Productos
            sCadProd = Trim(RDatos!nTipoCred)
            Do While Not RDatos.EOF
                sCadImp = sCadImp & oImpre.ImpreFormat(RDatos!nMonto, 8)
                RDatos.MoveNext
                If RDatos.EOF Then
                    Exit Do
                End If
                If sCadProd = Trim(RDatos!nTipoCred) Then
                    Exit Do
                End If
            Loop
            sCadImp = sCadImp & lnSaltoLinDoc
        Loop
        RDatos.Close
        Set RDatos = Nothing
        sCadImp = sCadImp & lnSaltoLinDoc
        RAnalista.MoveNext
    Loop
    RAnalista.Close
    Set RAnalista = Nothing
    ImprimirMetasAnalistas = sCadImp & Chr$(12)
    Exit Function

ErrorImprimirMetasAnalistas:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'ImprimeClientesRelacionadosResumen
'*** PEAC 20080412
Public Function ImprimeClientesRelacionadosResumen(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
                                ByVal psCodUsu As String, ByVal psNomCmac As String) As String

Dim sSql As String
Dim oDCredDoc As COMDCredito.DCOMCredito
Dim oDCred As COMDCredito.DCOMCredito
Dim oCredRel As COMDCredito.DCOMCredito
Dim oDGarantia As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String
Dim sEdadTiempoFunc As String
Dim oCred As COMDCredito.DCOMCredito
Dim R2 As ADODB.Recordset

Dim RRel As ADODB.Recordset
Dim RRelInt As ADODB.Recordset

Set oDCred = New COMDCredito.DCOMCredito
Set RRel = oDCred.RecuperaClientesRelacionados(psCtaCod)
Set oDCred = Nothing

Set oCred = New COMDCredito.DCOMCredito
Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
Set oCred = Nothing

Set oCredRel = New COMDCredito.DCOMCredito
Set RRelInt = oCredRel.RecuperaCreditosVigentesRelacCred(psCtaCod)
Set oCredRel = Nothing


'---------

    Set oDCredDoc = New COMDCredito.DCOMCredito
    'Set R = oDCredDoc.RecuperaDatosResumenComite(psCtaCod)
    Set R = oDCredDoc.RecuperaDatosClientesRelacionados(psCtaCod)
    Set oDCredDoc = Nothing
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing
    
    Set oDGarantia = New COMDCredito.DCOMGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
    sCadImp = ""
    If R.EOF Then Exit Function
      
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE" & IIf(R2.RecordCount > 0, " DE CREDITO REFINANCIADO", ""), psNomCmac, lsTab, False)
    Call ImprimeCabeceraDocumentoCred(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE" & IIf(R2.RecordCount > 0, " DE CREDITO REFINANCIADO", ""), psNomCmac, lsTab, False, "", "", R!codana, R!cPersNombre, R!cCtaCod, R!cTipoCred, R!cPersCod, R!DNI, R!RUC)
    
   
'*** Resumen clientes relacionados
    '--------------------------------------------------------------
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(30) & "DATOS PRINCIPALES DEL CREDITO" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "ANALISTA RESPONSABLE         : " & R!NomAna & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DESTINO DEL CREDITO          : " & R!cDestino & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "CONDICION DEL CREDITO        : " & R!cCondicion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "FECHA DE SOLICITUD           : " & R!FecSoli & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "VCTO. SOLICITADO A PARTIR DE : " & "....." & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "VCTO. SUGERIDO A PARTIR DE   : " & "....." & lnSaltoLinDoc & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "CODIGO DE LINEA  : " & R!cLineaCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & lnSaltoLinDoc
    
    If RGarantCred.EOF Then
        sCadImp = sCadImp & lsTab & "GARANTIA : " & oImpre.ImpreFormat(0, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(0, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(0, 15, 2) & lnSaltoLinDoc & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "GARANTIA : " & oImpre.ImpreFormat(RGarantCred!cTpoGarantia, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(RGarantCred!nPorGravar, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(RGarantCred!nGravado, 15, 2) & lnSaltoLinDoc & lnSaltoLinDoc
    End If
    
'------------------------------------------------------------------------------------
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 26, 0) & oImpre.ImpreFormat("MONTO", 18, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("NRO CUOTAS", 15, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 15, 2) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nPlazo = 0, 30, REstadosCred!nPlazo), 15, 0) & " dias" & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    
'--------------------------------------------------------------------------------------------------

    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(" CODIGO ", 15, 0) & oImpre.ImpreFormat("NOMBRE", 40, 0) & oImpre.ImpreFormat("RELACION", 20, 0) & oImpre.ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        'If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(RRelaCred!cPersCod, 15, 0) & oImpre.ImpreFormat(RRelaCred!cPersNombre, 40, 0) & oImpre.ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & oImpre.ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
        'End If
        RRelaCred.MoveNext
    Loop
    
'----------------------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "RELACIONADOS INTERNAS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(" CLIENTE ", 15, 0) & Space(22) & oImpre.ImpreFormat("RELACION DEL SOLCTE.", 26, 0) & oImpre.ImpreFormat("CREDITO", 20, 0) & oImpre.ImpreFormat("DIAS ATRASO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelInt.EOF
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(RRelInt!cPersNombre, 40, 0) & oImpre.ImpreFormat(RRelInt!cRelacion, 20, 0) & oImpre.ImpreFormat(RRelInt!cCtaCod, 18, 0) & oImpre.ImpreFormat(RRelInt!nDiasAtraso, 8, 0) & lnSaltoLinDoc
        RRelInt.MoveNext
    Loop
        
        
'----------------------------------------------------------------------------------------------------------------
    R.Close
    Set R = Nothing
    RRelaCred.Close
    Set RRelaCred = Nothing
    REstadosCred.Close
    Set REstadosCred = Nothing
    RGarantCred.Close
    Set RGarantCred = Nothing
    
    ImprimeClientesRelacionadosResumen = sCadImp
End Function

'*** PEAC 20080412
Public Function ImprimeClientesRelacionados(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
                                ByVal psCodUsu As String, ByVal psNomCmac As String, ByVal pDocNat As String, ByVal pDocJur As String) As String

Dim sSql As String
Dim oDCredDoc As COMDCredito.DCOMCredito
Dim oDCred As COMDCredito.DCOMCredito
Dim oDGarantia As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String
Dim sEdadTiempoFunc As String
Dim oCred As COMDCredito.DCOMCredito
Dim oCredRel As COMDCredito.DCOMCredito
Dim R2 As ADODB.Recordset
Dim RCaliSbs As ADODB.Recordset

Dim RRel As ADODB.Recordset
Dim RRelInt As ADODB.Recordset

'*** PEAC 20160820
Dim oCredFormEval As COMDCredito.DCOMFormatosEval
Set oCredFormEval = New COMDCredito.DCOMFormatosEval
Dim rsFormEval As ADODB.Recordset
Set rsFormEval = oCredFormEval.ObtieneDatosComunes(psCtaCod)
Set oCredFormEval = Nothing
'*** FIN PEAC
'-------------------------------------------------------------------
    Set oDCred = New COMDCredito.DCOMCredito
    Set RRel = oDCred.RecuperaClientesRelacionados(psCtaCod)
    Set oDCred = Nothing

    Set oCred = New COMDCredito.DCOMCredito
    Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
    Set oCred = Nothing

    Set oDCred = New COMDCredito.DCOMCredito
    Set RCaliSbs = oDCred.RecuperaCaliSbs(pDocNat, pDocJur)
    Set oDCred = Nothing

    Set oDCredDoc = New COMDCredito.DCOMCredito
    Set R = oDCredDoc.RecuperaDatosClientesRelacionados(psCtaCod)
    Set oDCredDoc = Nothing

    Set oDCred = New COMDCredito.DCOMCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing

    Set oDGarantia = New COMDCredito.DCOMGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
'-------------------------------------------------------------------------
    
    Set oCredRel = New COMDCredito.DCOMCredito
    Set RRelInt = oCredRel.RecuperaCreditosVigentesRelacCred(psCtaCod)
    Set oCredRel = Nothing

'-------------------------------------------------------------------------
    sCadImp = ""
    If R.EOF Then Exit Function
      
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE" & IIf(R2.RecordCount > 0, " DE CREDITO REFINANCIADO", ""), psNomCmac, lsTab, False)
    Call ImprimeCabeceraDocumentoCred(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE" & IIf(R2.RecordCount > 0, " DE CREDITO REFINANCIADO", ""), psNomCmac, lsTab, False, "", "", R!codana, R!cPersNombre, R!cCtaCod, R!cTipoCred, R!cPersCod, R!DNI, R!RUC)
    
    '--------------------------------------------------------------

    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(30) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "COD. SBS : " & R!cperscodsbs & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "CALIFICACION: " & "  0: " & RCaliSbs!nNormal & "  1: " & RCaliSbs!nPotencial & "  2: " & RCaliSbs!nDeficiente & "  3: " & RCaliSbs!nDudoso & "  4: " & RCaliSbs!nPerdido & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "FECHA INSCRIPCION: " & IIf(R!dPersIngRuc <= "01/01/1950", "  /  /    ", R!dPersIngRuc) & " FECHA INI.ACTIVIDAD/INGRESO: " & IIf(R!dPersIniActi <= "01/01/1950", "  /  /    ", R!dPersIniActi) & "EXPERIENCIA SECTOR : " & R!nExperiencia & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DIRECCION: " & R!cPersDireccDomicilio & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "CODIGO CIIU: " & R!cPersCIIU & " " & R!cCiiuDescripcion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TELEFONO: " & R!telf & " CELULAR: " & R!Celular & " E-MAIL: " & R!Email & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DIRECCION NEGOCIO/LABORAL: " & R!DireTrabajo & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TELEFONO: " & R!TelfTrabajo & " CELULAR : " & "....." & " E-MAIL : " & "....." & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "FECHA DE NACIMIENTO: " & R!dPersNacCreac & " SEXO : " & R!Sexo & " NACIONALIDAD : " & R!Nacionalidad & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "EDAD : " & R!edad & " ESTADO CIVIL: " & R!cEstadoCivil & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Nº DE DEPENDIENTES-INCLUIR CONYUGUE: " & oImpre.ImpreFormat(R!nNumDependi, 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "NIVEL DE INSTRUCCION: " & "...." & " SITUACION LABORAL : " & "....." & lnSaltoLinDoc
'    sCadImp = sCadImp & lsTab & "FUENTE ING.DEPENDIENTE: " & IIf(R!TipoFuente = "D", R!empresa, "----------") & " GIRO O ACTIVIDAD: " & R!cActiGiro & lnSaltoLinDoc
    If Trim(rsFormEval!nCodForm) = 0 Then
        sCadImp = sCadImp & lsTab & "FUENTE ING.DEPENDIENTE: " & IIf(R!TipoFuente = "D", R!empresa, "----------") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "GIRO O ACTIVIDAD: " & R!cActiGiro & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "FORMATO DE EVALUACION: " & Trim(rsFormEval!cformato) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "GIRO O ACTIVIDAD: " & Trim(rsFormEval!cActiGiro) & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc

    Do While Not RRel.EOF
        sCadImp = sCadImp & lsTab & "DATOS DEL " & RRel!Relacion & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc

        sCadImp = sCadImp & lsTab & "COD. CLIENTE : " & RRel!cPersCod & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "NOMBRE : " & RRel!cPersNombre & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "DOC NATUTAL : " & RRel!docnatu & "DOC. JURIDICO : " & RRel!docJuri & "COD. SBS : " & RRel!codsbs & lnSaltoLinDoc
        If RRel!cNatJur = "1" Then
            sCadImp = sCadImp & lsTab & "CALIFICACION : " & "  0: " & RRel!nNormal1 & "  1: " & RRel!nPotencial1 & "  2: " & RRel!nDeficiente1 & "  3: " & RRel!nDudoso1 & "  4: " & RRel!nPerdido1 & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lsTab & "CALIFICACION : " & "  0: " & RRel!nNormal2 & "  1: " & RRel!nPotencial2 & "  2: " & RRel!nDeficiente2 & "  3: " & RRel!nDudoso2 & "  4: " & RRel!nPerdido2 & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & "DIRECCION : " & RRel!dire & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "TELEFONO : " & RRel!telf & " CELULAR :" & RRel!Celular & " E-MAIL :" & RRel!Email & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "DIRECCION NEGOCIO/LABORAL : " & RRel!DireTrabajo & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "TELEFONO : " & RRel!TelfTrabajo & " CELULAR :" & "....." & " E-MAIL :" & "....." & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FECHA DE NACIMIENTO : " & RRel!FecNaci & " SEXO :" & RRel!Sexo & " NACIONALIDAD :" & RRel!Nacionalidad & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "EDAD : " & RRel!edad & " ESTADO CIVIL :" & RRel!EstadoCivil & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "NIVEL DE INSTRUCCION : " & "....." & " SITUACION LABORAL : " & "....." & lnSaltoLinDoc
        'sCadImp = sCadImp & lsTab & "FUENTE ING.DEPENDIENTE: " & RRel!empresa & " GIRO O ACTIVIDAD :" & RRel!cActiGiro & lnSaltoLinDoc
        'sCadImp = sCadImp & lsTab & "FORMATO DE EVALUACION: " & Trim(rsFormEval!cformato) & " GIRO O ACTIVIDAD: " & Trim(rsFormEval!cActiGiro) & lnSaltoLinDoc 'LUCV20170727, Comentó según observación RUSI
       
        If Trim(rsFormEval!nCodForm) = 0 Then
            sCadImp = sCadImp & lsTab & "FUENTE ING.DEPENDIENTE: " & RRel!empresa & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "GIRO O ACTIVIDAD: " & Trim(RRel!cActiGiro) & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lsTab & "GIRO O ACTIVIDAD: " & Trim(RRel!cActiGiro) & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        
        'sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 15, 2) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nPlazo = 0, 30, REstadosCred!nPlazo), 15, 0) & " dias" & lnSaltoLinDoc
        RRel.MoveNext
    Loop
'--------------------------------------------------------OTRO REPORTE -------------------------------------------------------------------

    sCadImp = sCadImp & lnSaltoPagina
    
   
'*** Resumen clientes relacionados
    '--------------------------------
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & Space(30) & "DATOS PRINCIPALES DEL ARRENDAMIENTO FINANCIERO" & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & Space(30) & "DATOS PRINCIPALES DEL CREDITO" & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "ANALISTA RESPONSABLE         : " & R!NomAna & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & "DESTINO DE LA OPERACION      : " & R!cDestino & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "CONDICION DE LA OPERACION    : " & R!cCondicion & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "DESTINO DEL CREDITO          : " & R!cDestino & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "CONDICION DEL CREDITO        : " & R!cCondicion & lnSaltoLinDoc
    End If
    
    sCadImp = sCadImp & lsTab & "FECHA DE SOLICITUD           : " & R!FecSoli & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & "VCTO. SOLICITADO A PARTIR DE : " & "....." & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "VCTO. SUGERIDO A PARTIR DE   : " & IIf(Year(R!FecSuge) <= "1950", "  /  /    ", Format(R!FecSuge, "dd/mm/yyyy")) & lnSaltoLinDoc & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "CODIGO DE LINEA  : " & R!cLineaCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & lnSaltoLinDoc
    
    If RGarantCred.EOF Then
        sCadImp = sCadImp & lsTab & "GARANTIA : " & oImpre.ImpreFormat(0, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(0, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(0, 15, 2) & lnSaltoLinDoc & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "GARANTIA : " & oImpre.ImpreFormat(RGarantCred!cTpoGarantia, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(RGarantCred!nPorGravar, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(RGarantCred!nGravado, 15, 2) & lnSaltoLinDoc & lnSaltoLinDoc
    End If
    
'--------------------------------
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 26, 0) & oImpre.ImpreFormat("MONTO", 18, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("NRO CUOTAS", 15, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 15, 2) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nPlazo = 0, 30, REstadosCred!nPlazo), 15, 0) & " dias" & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    
'---------------------------------

    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS - RESUMEN" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(" CODIGO ", 15, 0) & oImpre.ImpreFormat("NOMBRE", 40, 0) & oImpre.ImpreFormat("RELACION", 20, 0) & oImpre.ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        'If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(RRelaCred!cPersCod, 15, 0) & oImpre.ImpreFormat(RRelaCred!cPersNombre, 40, 0) & oImpre.ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & oImpre.ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
        'End If
        RRelaCred.MoveNext
    Loop
    
'---------------------------------

    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "RELACIONADOS INTERNAS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(" CLIENTE ", 15, 0) & Space(22) & oImpre.ImpreFormat("RELACION DEL SOLCTE.", 26, 0) & oImpre.ImpreFormat("CREDITO", 20, 0) & oImpre.ImpreFormat("DIAS ATRASO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelInt.EOF
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(RRelInt!cPersNombre, 40, 0) & oImpre.ImpreFormat(RRelInt!cRelacion, 20, 0) & oImpre.ImpreFormat(RRelInt!cCtaCod, 18, 0) & oImpre.ImpreFormat(RRelInt!nDiasAtraso, 8, 0) & lnSaltoLinDoc
        RRelInt.MoveNext
    Loop
    
'----------------------------------------------------------------------------------------------------------------
        
    R.Close
    Set R = Nothing
    RRelaCred.Close
    Set RRelaCred = Nothing
    REstadosCred.Close
    Set REstadosCred = Nothing
    RGarantCred.Close
    Set RGarantCred = Nothing
    rsFormEval.Close
    Set rsFormEval = Nothing
    
    ImprimeClientesRelacionados = sCadImp
End Function



Public Function ImprimeResumenComite(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
                                ByVal psCodUsu As String, ByVal psNomCmac As String) As String
Dim sSql As String
Dim oDCredDoc As COMDCredito.DCOMCredDoc
Dim oDCred As COMDCredito.DCOMCredito
Dim oDGarantia As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String

Dim sEdadTiempoFunc As String

'ARCV 29-01-2007
Dim oCred As COMDCredito.DCOMCredito
Dim R2 As ADODB.Recordset

Set oCred = New COMDCredito.DCOMCredito
Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
Set oCred = Nothing
'---------

    Set oDCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oDCredDoc.RecuperaDatosResumenComite(psCtaCod)
    Set oDCredDoc = Nothing
    Set oDCred = New COMDCredito.DCOMCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing
    Set oDGarantia = New COMDCredito.DCOMGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
    sCadImp = ""
    If R.EOF Then Exit Function
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE" & IIf(R2.RecordCount > 0, " DE CREDITO REFINANCIADO", ""), psNomCmac, lsTab, False)
    sCadImp = sCadImp & lnSaltoLinDoc
    'ALPA 20120424
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Operación AF: " & psCtaCod & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), 5, 5) & Space(25) & R!cTipoCred & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & oImpre.ImpreFormat(IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), 5, 5) & Space(25) & R!cTipoCred & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(2) & oCad.PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   : " & R!DNI & Space(30) & "Doc. Juridico : " & R!RUC & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    
    If R!nPersPersoneria <> 1 Then
       sEdadTiempoFunc = "Tiempo de Func.: "
    Else
       sEdadTiempoFunc = "Edad           : "
    End If
    
    If IsNull(R!dPersNacCreac) Then
        sCadImp = sCadImp & lsTab & sEdadTiempoFunc
    Else
        sCadImp = sCadImp & lsTab & sEdadTiempoFunc & Format(Year(CDate(psFecSis)) - Year(R!dPersNacCreac), "#0") & " Años "
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    If R!nPersPersoneria <> 1 Then
       sCadImp = sCadImp & Space(45) & lnSaltoLinDoc
    Else
       sCadImp = sCadImp & Space(30) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    If R!nPersPersoneria <> 1 Then
       sCadImp = sCadImp & lnSaltoLinDoc
    Else
       sCadImp = sCadImp & lsTab & "Nro Hijos : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(" CODIGO ", 15, 0) & oImpre.ImpreFormat("NOMBRE", 40, 0) & oImpre.ImpreFormat("RELACION", 20, 0) & oImpre.ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(RRelaCred!cPersCod, 15, 0) & oImpre.ImpreFormat(RRelaCred!cPersNombre, 40, 0) & oImpre.ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & oImpre.ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
        End If
        RRelaCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & Space(30) & "DATOS DEL ARRENDAMIENTO FINANCIERO" & lnSaltoLinDoc & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & Space(30) & "DATOS DEL CREDITO" & lnSaltoLinDoc & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("ESTADO", 26, 0) & oImpre.ImpreFormat("MONTO", 18, 0) & oImpre.ImpreFormat("CUOTA", 15, 0) & oImpre.ImpreFormat("NRO CUOTAS", 15, 0) & oImpre.ImpreFormat("PLAZO", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(REstadosCred!cEstado, 15, 0) & oImpre.ImpreFormat(REstadosCred!nMonto, 15, 2) & oImpre.ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & oImpre.ImpreFormat(REstadosCred!nCuotas, 15, 0) & oImpre.ImpreFormat(IIf(REstadosCred!nPlazo = 0, 30, REstadosCred!nPlazo), 15, 0) & " dias" & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & "DESTINO : " & R!cDestino & Space(20) & "LINEA DE CREDITO : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & lnSaltoLinDoc
    
    'ARCV 01-02-2007
    If RGarantCred.EOF Then
        sCadImp = sCadImp & "GARANTIA : " & oImpre.ImpreFormat(0, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(0, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(0, 15, 2) & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & "GARANTIA : " & oImpre.ImpreFormat(RGarantCred!cTpoGarantia, 20, 0) & " Garant. Total: " & oImpre.ImpreFormat(RGarantCred!nPorGravar, 10, 2) & " Garant. Cred. : " & oImpre.ImpreFormat(RGarantCred!nGravado, 15, 2) & lnSaltoLinDoc
    End If
    '----------
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & Space(35) & "RESOLUCION DE COMITE" & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & Space(15) & "MONTO" & Space(5) & "# CUOTAS " & Space(10) & "MONTO CUOTA" & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APROBADO : " & ".... ......................    .......................      ....................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Tipo de Contrato :      Fondo : ................  Garantias : ................................" & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(1) & "____________________________      ______________________________    _________________________" & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    
    'ARCV 30-01-2007
    sCadImp = sCadImp & lsTab & "............................................................................................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "............................................................................................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "............................................................................................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "............................................................................................." & lnSaltoLinDoc & lnSaltoLinDoc
    '---------------
    sCadImp = sCadImp & Chr$(12)
    
    R.Close
    Set R = Nothing
    RRelaCred.Close
    Set RRelaCred = Nothing
    REstadosCred.Close
    Set REstadosCred = Nothing
    RGarantCred.Close
    Set RGarantCred = Nothing
    
    ImprimeResumenComite = sCadImp
End Function

Public Function ImprimePlandePagosDuplicado(ByVal psCtaCod As String, ByVal psNomAge As String, _
                        ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
                        ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
                        Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0, _
                        Optional ByVal pnMivivienda As Integer, Optional ByVal pbDetallado As Boolean = False, _
                        Optional ByVal psFechaDes As String) As String
'*****************Parametro psFecDes agregado por PASI20131128 TIERS136-2013
    
Dim lsCadImp As String
Dim lsCadImp_2 As String

'****************************************************
'Modificado por PASI20131128 TIERS136-2013
'lsCadImp = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, , psNomCmac, pbComodin, pbCalendDin, pbDetallado)
'If pnMivivienda Then
'    lsCadImp_2 = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, True, psNomCmac, pbComodin, pbCalendDin, pbDetallado)
'    lsCadImp = lsCadImp & lsCadImp_2
'End If

'lsCadImp = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, , psNomCmac, pbComodin, pbCalendDin, pbDetallado, , psFechaDes)
lsCadImp = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, , psNomCmac, pbComodin, pbCalendDin, pbDetallado, , psFechaDes, True) 'JUEZ 20151229

'WIOR 20160112 ***
Dim oCred As COMDCredito.DCOMCredito
Dim rsCred As ADODB.Recordset
Dim bMIVIVIENDAAnt As Boolean

Set oCred = New COMDCredito.DCOMCredito

Call oCred.CargaDatosReportesDuplicados(psCtaCod, Nothing, rsCred, "")

bMIVIVIENDAAnt = oCred.EsCredMIVIVENDA(rsCred!cTpoProdCod, rsCred!cTpoCredCod, 1)
'WIOR FIN ********

'If pnMivivienda Then'WIOR 20160112 - COMENTO
If bMIVIVIENDAAnt Then 'WIOR 20160112
    'lsCadImp_2 = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, True, psNomCmac, pbComodin, pbCalendDin, pbDetallado, , psFechaDes)
    lsCadImp_2 = ImprimePlandePagos(psCtaCod, psNomAge, Format(psFecSis, "dd/mm/yyyy"), psCodUsu, pnMontoPrestamo, pbMiViv, True, psNomCmac, pbComodin, pbCalendDin, pbDetallado, , psFechaDes, True) 'JUEZ 20151229
    lsCadImp = lsCadImp & lsCadImp_2
End If
'End PASI********************************************

ImprimePlandePagosDuplicado = lsCadImp

End Function

Public Function ImprimeRegistroSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
                    ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psTipoPers As String, _
                    ByVal psCodCli As String, ByVal psNombreCli As String, ByVal psDocNat As String, _
                    ByVal psDocJur As String, ByVal psDirecc As String, _
                    ByVal psNroHijos As String, ByVal psEdad As String, ByVal psEstadoCiv As String, _
                    ByVal psConyugue As String, ByVal psLECony As String, ByVal psRazonSoc As String, _
                    ByVal psTipoFuente As String, ByVal psRazonSocDirecc As String, _
                    ByVal psActividad As String, ByVal psIniAct As String, ByVal MatRelaCli As Variant, _
                    ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, ByVal psLineaCred As String, _
                    ByVal psFechaSol As String, ByVal MatDatosEstado As Variant, ByVal MatGarantias As Variant, _
                    ByVal MatRelacInterna As Variant, ByVal psNomCmac As String) As String
    
Dim sCadImp As String
Dim i As Integer

    sCadImp = ""
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "REGISTRO DE SOLICITUD DE ARRENDAMIENTO FINANCIERO", psNomCmac, Space(5), False)
    Else
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "REGISTRO DE SOLICITUD DE CREDITO", psNomCmac, Space(5), False)
    End If
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & "Operación AF : " & psCtaCod & Space(40) & psTipoCred & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & "Credito : " & psCtaCod & Space(40) & psTipoCred & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(40) & psTipoPers & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        :" & psCodCli & Space(5) & oCad.PstaNombre(psNombreCli) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   :" & psDocNat & Space(15) & "Doc. Juridico :" & psDocJur & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      :" & psDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Edad           :" & psEdad & " Años" & Space(30) & "Estado Civil : " & psEstadoCiv & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Conyugue       :" & oImpre.ImpreFormat(psConyugue, 20, 2, 0) & " L.E:" & psLECony & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE INGRESO" & Space(30) & psTipoFuente & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & "FORMATO DE EVALUACION" & Space(30) & psTipoFuente & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Razon Social           :" & psRazonSoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion              :" & psRazonSocDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Actividad              :" & psActividad & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Inicio de Actividades  :" & psIniAct & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("CODIGO", 10, 0) & oImpre.ImpreFormat("NOMBRE", 30, 0) & oImpre.ImpreFormat("RELACION", 12, 0) & oImpre.ImpreFormat("DOC.IDENT.", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    For i = 0 To UBound(MatRelaCli) - 1
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(MatRelaCli(i, 0), 10, 0) & oImpre.ImpreFormat(MatRelaCli(i, 1), 30, 0) & oImpre.ImpreFormat(MatRelaCli(i, 2), 12, 0) & oImpre.ImpreFormat(MatRelaCli(i, 3), 15, 0) & lnSaltoLinDoc
    Next i
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "DATOS DEL ARRENDAMIENTO FINANCIERO" & lsNegritaOff & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "DATOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & lsTab & "Analista : " & psAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Destino : " & oImpre.ImpreFormat(psDestino, 20, 0) & "Condicion : " & psCondicion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Linea Credito : " & oImpre.ImpreFormat(psLineaCred, 20, 0) & "Fecha de Solicitud : " & psFechaSol & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & oImpre.ImpreFormat("Estado", 20, 0) & oImpre.ImpreFormat("Monto", 16, 0) & oImpre.ImpreFormat("Cuota", 16, 0) & oImpre.ImpreFormat("Cuotas", 15, 0) & oImpre.ImpreFormat("Plazo", 10, 0) & oImpre.ImpreFormat("Superavit", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    For i = 0 To 1
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(MatDatosEstado(i, 0), 15) & oImpre.ImpreFormat(MatDatosEstado(i, 1), 15) & oImpre.ImpreFormat(MatDatosEstado(i, 2), 15) & oImpre.ImpreFormat(MatDatosEstado(i, 3), 15) & oImpre.ImpreFormat(MatDatosEstado(i, 4), 7) & oImpre.ImpreFormat(MatDatosEstado(i, 5), 15) & lnSaltoLinDoc
    Next i
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "GARANTIAS" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("Garantia", 30, 0) & oImpre.ImpreFormat("Total", 20, 0) & oImpre.ImpreFormat("Gravado", 20, 0) & lnSaltoLinDoc
    For i = 0 To UBound(MatGarantias) - 1
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(MatGarantias(i, 0), 30, 0) & oImpre.ImpreFormat(MatGarantias(i, 1) & " " & MatGarantias(i, 2), 20, 0) & oImpre.ImpreFormat(MatGarantias(i, 3) & " " & MatGarantias(i, 4), 20, 0) & lnSaltoLinDoc
    Next i
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(30) & lsNegritaOn & "RELACIONES INTERNAS (Creditos Vigentes)" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("CLIENTE", 30, 0) & oImpre.ImpreFormat("RELACION", 15, 0) & oImpre.ImpreFormat("CREDITO", 20, 0) & oImpre.ImpreFormat("ATRASO", 15, 0) & oImpre.ImpreFormat("SALDO K.", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    If IsArray(MatRelacInterna) Then
        For i = 0 To UBound(MatRelacInterna) - 1
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(MatRelacInterna(i, 0), 30, 0) & oImpre.ImpreFormat(MatRelacInterna(i, 1), 15, 0) & oImpre.ImpreFormat(MatRelacInterna(i, 2), 20, 0) & oImpre.ImpreFormat(MatRelacInterna(i, 3), 15, 0) & oImpre.ImpreFormat(MatRelacInterna(i, 4), 10, 2, True) & lnSaltoLinDoc
        Next i
    Else
        sCadImp = sCadImp & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & Chr$(12)
    ImprimeRegistroSolicitud = sCadImp
End Function


Public Function ImprimeRegistroSolicitudDuplicado(ByVal psCtaCod As String, _
                                                ByVal psNomAge As String, _
                                                ByVal psFecha As String, _
                                                ByVal psCodUsu As String, _
                                                ByVal psNomCmac As String) As String

Dim oPersona As COMDPersona.DCOMPersona
Dim oPersConyugue As COMDPersona.DCOMPersona
Dim oPersRelac As COMDPersona.UCOMPersona
Dim oDCred As COMDCredito.DCOMCredito
Dim RRelacCred As ADODB.Recordset
Dim RDatosCom As ADODB.Recordset
Dim RFteIngreso As ADODB.Recordset
Dim RColocEstado As ADODB.Recordset
Dim RRelacInterna As ADODB.Recordset
Dim psTitular As String
Dim sCodConyugue As String
Dim psCodPersRazon As String
Dim oCons As COMDConstantes.DCOMConstantes
Dim sCadImp As String
Dim oDDatosFormEval As COMDCredito.DCOMFormatosEval
Set oDDatosFormEval = New COMDCredito.DCOMFormatosEval
Dim rsDatFormEval As ADODB.Recordset

            Set oDCred = New COMDCredito.DCOMCredito
            Set RRelacCred = oDCred.RecuperaRelacPers(psCtaCod)
            Set RDatosCom = oDCred.RecuperaDatosComunes(psCtaCod)
            Set RFteIngreso = oDCred.RecuperaColocFteIngreso(psCtaCod)
            Set RColocEstado = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
            Set RRelacInterna = oDCred.RecuperaCreditosVigentesRelacCred(psCtaCod)
            Set oDCred = Nothing
            psTitular = ""
            sCodConyugue = ""
            Do While Not RRelacCred.EOF
                If RRelacCred!nConsValor = gColRelPersTitular Then
                    psTitular = RRelacCred!cPersCod
                End If
                If RRelacCred!nConsValor = gColRelPersConyugue Then
                    sCodConyugue = RRelacCred!cPersCod
                End If
                RRelacCred.MoveNext
            Loop
            If psTitular <> "" Then
                Set oPersona = New COMDPersona.DCOMPersona
                Call oPersona.RecuperaPersona(psTitular)
            End If
            
            Dim sNomCony As String
            Dim sDNIConyugue As String
            If sCodConyugue <> "" Then
                Set oPersConyugue = New COMDPersona.DCOMPersona
                Call oPersConyugue.RecuperaPersona(sCodConyugue)
                sNomCony = oPersConyugue.NombreCompleto
                sDNIConyugue = oPersConyugue.ObtenerDNI
                Set oPersConyugue = Nothing
            Else
                sNomCony = ""
                sDNIConyugue = ""
            End If
            
            Dim sTipoFte As String
            Dim nIngresoFam As Double
            Dim sRazonSoc As String
            Dim sFteDirecc As String
            Dim sActividad As String
            Dim dFecIniAct As Date
            Dim oPersFteIng As New COMDPersona.DCOMPersona
            Dim nPunt As Integer
            'nPunt = oPersona.UbicarFteIngreso(RFteIngreso!cPersFIPersCod)
            
            '*** PEAC 20160820
'            If nPunt <> -1 Then
'                sTipoFte = oPersona.ObtenerFteIngTipo(nPunt)
'                If Len(Trim(sTipoFte)) > 0 Then
'                    Set oCons = New COMDConstantes.DCOMConstantes
'                    sTipoFte = oCons.DameDescripcionConstante(gPersFteIngresoTipo, CInt(sTipoFte))
'                    Set oCons = Nothing
'                End If
'                sRazonSoc = oPersona.ObtenerFteIngRazonSocial(nPunt)
'                Set oPersFteIng = New COMDPersona.DCOMPersona
'                Call oPersFteIng.RecuperaPersona(oPersona.ObtenerFteIngFuente(nPunt))
'                sFteDirecc = oPersFteIng.Domicilio
'                sActividad = oPersFteIng.CIIUDescripcion
'                dFecIniAct = oPersFteIng.FechaNacimiento
'                'nIngresoFam = Format(oPersFteIng.ObtenerFteIngIngresosTotales(nPunt), "#0.00")
'                nIngresoFam = 0
'                Set oPersFteIng = Nothing
'            Else
'                sTipoFte = ""
'                sRazonSoc = ""
'                sFteDirecc = ""
'                nIngresoFam = 0
'            End If
                Set rsDatFormEval = oDDatosFormEval.ObtieneDatosComunes(psCtaCod)
                sTipoFte = rsDatFormEval!cformato
                sRazonSoc = rsDatFormEval!cPersNombre
                sFteDirecc = rsDatFormEval!cPersNegocioDireccion
                nIngresoFam = 0
                sActividad = rsDatFormEval!cActiGiro
            '*** FIN PEAC
            
            'Carga Relaciones de Cliente
            Dim MatRelaCli() As String
            ReDim MatRelaCli(RRelacCred.RecordCount, 4)
            nPunt = 0
            Do While Not RRelacCred.EOF
                If RRelacCred!nConsValor <> gColRelPersTitular Then
                    MatRelaCli(nPunt, 0) = RRelacCred!cPersCod
                    MatRelaCli(nPunt, 1) = RRelacCred!cPersNombre
                    MatRelaCli(nPunt, 2) = RRelacCred!cConsDescripcion
                    MatRelaCli(nPunt, 4) = RRelacCred!DNI
                    nPunt = nPunt + 1
                End If
                RRelacCred.MoveNext
            Loop
            RRelacCred.Close
            Set RRelacCred = Nothing
            
            'Carga Estados del Credito
            Dim MatEstados(2, 6) As String
            Do While Not RColocEstado.EOF
                Select Case RColocEstado!nPrdEstado
                    Case gColocEstSolic
                        MatEstados(0, 0) = "Solicitado"
                        MatEstados(0, 1) = Format(RColocEstado!nMonto, "#0.00")
                        MatEstados(0, 2) = ""
                        MatEstados(0, 3) = Trim(Str(RColocEstado!nCuotas))
                        MatEstados(0, 4) = Trim(Str(RColocEstado!nPlazo))
                        MatEstados(0, 5) = ""
                    Case gColocEstSug
                        MatEstados(1, 0) = "Sugerido"
                        MatEstados(1, 1) = Format(IIf(IsNull(RColocEstado!nMonto), 0, RColocEstado!nMonto), "#0.00")
                        MatEstados(1, 2) = Format(IIf(IsNull(RColocEstado!nMontoCuota), 0, RColocEstado!nMontoCuota), "#0.00")
                        MatEstados(1, 3) = Trim(Str(IIf(IsNull(RColocEstado!nCuotas), 0, RColocEstado!nCuotas)))
                        MatEstados(1, 4) = Trim(Str(IIf(IsNull(RColocEstado!nPlazo), 0, RColocEstado!nPlazo)))
                        MatEstados(1, 5) = Format(nIngresoFam - CDbl(MatEstados(1, 2)), "#0.00")
                        Exit Do
                End Select
                RColocEstado.MoveNext
            Loop
            RColocEstado.Close
            Set RColocEstado = Nothing
            
            'Carga las Garantias
            Dim MatGarantias() As String
            Dim RGarantias As ADODB.Recordset
            Dim oGarantia As COMDCredito.DCOMGarantia
            Set oGarantia = New COMDCredito.DCOMGarantia
            Set RGarantias = oGarantia.RecuperaGarantiaCredito(psCtaCod)
            Set oGarantia = Nothing
            ReDim MatGarantias(RGarantias.RecordCount, 5)
            Do While Not RGarantias.EOF
                MatGarantias(RGarantias.Bookmark - 1, 0) = RGarantias!cTpoGarantia
                MatGarantias(RGarantias.Bookmark - 1, 1) = Format(RGarantias!nGravado, "#0.00")
                MatGarantias(RGarantias.Bookmark - 1, 2) = RGarantias!cMoneda
                MatGarantias(RGarantias.Bookmark - 1, 3) = Format(RGarantias!nGravament, "#0.00")
                MatGarantias(RGarantias.Bookmark - 1, 4) = RGarantias!cMonedaGar
                RGarantias.MoveNext
            Loop
            RGarantias.Close
            Set RGarantias = Nothing
            
            'Carga los Creditos Vigentes que tienen las personas relacionadas con el Credito
            Dim MatRelaInterna() As String
            If RRelacInterna.RecordCount > 0 Then
                ReDim MatRelaInterna(RRelacInterna.RecordCount, 5)
                Do While Not RRelacInterna.EOF
                    MatRelaInterna(RRelacInterna.Bookmark - 1, 0) = oCad.PstaNombre(RRelacInterna!cPersNombre)
                    MatRelaInterna(RRelacInterna.Bookmark - 1, 1) = Trim(RRelacInterna!cRelacion)
                    MatRelaInterna(RRelacInterna.Bookmark - 1, 2) = RRelacInterna!cCtaCod
                    MatRelaInterna(RRelacInterna.Bookmark - 1, 3) = Trim(Str(RRelacInterna!nDiasAtraso))
                    MatRelaInterna(RRelacInterna.Bookmark - 1, 4) = Trim(Format(RRelacInterna!nSaldo, "#,#.00"))
                    RRelacInterna.MoveNext
                Loop
            Else
                ReDim MatRelaInterna(0, 0)
            End If
            RRelacInterna.Close
            Set RRelacInterna = Nothing
            Set oCons = New COMDConstantes.DCOMConstantes
            Dim sEstadoCivil As Long
            
            If oPersona.Personeria > 1 Then
                sEstadoCivil = 0
            Else
                sEstadoCivil = oPersona.EstadoCivil
            End If
            
            sCadImp = ImprimeRegistroSolicitud(psCtaCod, psNomAge, psFecha, psCodUsu, _
                IIf(IsNull(RDatosCom!cTipoCredDescrip), RDatosCom!cTpoCredDes, RDatosCom!cTipoCredDescrip), oCons.DameDescripcionConstante(gPersPersoneria, oPersona.Personeria), _
                oPersona.PersCodigo, oPersona.NombreCompleto, oPersona.ObtenerDNI, oPersona.ObtenerRUC, _
                oPersona.Domicilio, oPersona.Hijos, Format(Year(psFecha) - Year(oPersona.FechaNacimiento), "#0"), oCons.DameDescripcionConstante(gPersEstadoCivil, sEstadoCivil), _
                sNomCony, sDNIConyugue, sRazonSoc, sTipoFte, sFteDirecc, _
                sActividad, IIf(dFecIniAct = #12:00:00 AM#, "", Format(dFecIniAct, "dd/mm/yyyy")), MatRelaCli, RDatosCom!cAnalista, RDatosCom!cCondicion, _
                RDatosCom!cDestinoDescripcion, RDatosCom!cLineaCred, RDatosCom!dFechaSol, MatEstados, MatGarantias, MatRelaInterna, psNomCmac)
            
            Set oCons = Nothing

            ImprimeRegistroSolicitudDuplicado = sCadImp
            
End Function

Public Function ImprimeCalendarioDesemb(ByVal psCtaCod As String, ByVal psNomAge As String, _
                        ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
                        ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
                        Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0) As String

Dim R As ADODB.Recordset
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim oCalend As COMDCredito.DCOMCalendario
Dim oNCred As COMNCredito.NCOMCredito
Dim MatCalend As Variant
Dim sCadImp As String
Dim nCapital As Double
Dim RDesPar As ADODB.Recordset

    Set oCredDoc = New COMDCredito.DCOMCredDoc
    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
    Set oCredDoc = Nothing
    
   
    Set oCalend = New COMDCredito.DCOMCalendario
    Set RDesPar = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    Set oNCred = New COMNCredito.NCOMCredito
    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod, pnMontoPrestamo, pbCalendParalelo, IIf(pbCalendParalelo, True, False))
    Set oNCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
    sCadImp = sCadImp & lsTab & "Credito       : " & oImpre.ImpreFormat(psCtaCod, 30, 0) & Space(1) & "Agencia : " & psNomAge & Chr$(10)
    sCadImp = sCadImp & lsTab & "Cliente       : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
    'sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & Space(1) & "Cuota    : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tasa Interes  : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & Space(1) & "Plazo    : " & Trim(Str(R!nPlazo)) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Mensual" & Chr$(10)
    
'    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & oCad.PstaNombre(R!cPersNombre) & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Analista      : " & oCad.PstaNombre(R!cAnalista) & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & oImpre.ImpreFormat(R!cTipoCuota, 30, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
'    sCadImp = sCadImp & lsTab & "Interes       : " & oImpre.ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
'
    sCadImp = sCadImp & lsTab & "Moneda        : " & oImpre.ImpreFormat(R!cMoneda, 30, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & Chr$(10)

    sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & oImpre.ImpreFormat("Normal", 30, 0)
    
    sCadImp = sCadImp & lsTab & "Calend. Din.  :" & oImpre.ImpreFormat("NO", 30, 0)

    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    
    If RDesPar.RecordCount >= 2 Then
        sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("DESEMBOLSO", 10, 0) & oImpre.ImpreFormat("ESTADO", 10, 2) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        Do While Not RDesPar.EOF
            sCadImp = sCadImp & lsTab & oImpre.ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & oImpre.ImpreFormat(RDesPar!nCapital, 10, 2, True) & oImpre.ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
            RDesPar.MoveNext
        Loop
    End If
    RDesPar.Close
    
    sCadImp = sCadImp & Chr$(10)
    
   
    sCadImp = sCadImp & Chr$(27) & Chr$(84) & Chr$(1)
    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(1)
    
    sCadImp = sCadImp & Chr$(12)
    ImprimeCalendarioDesemb = sCadImp
End Function

Public Function ImprimeBoletaExtorno(ByVal psDescrip As String, ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC, _
            Optional ByVal psCodUserOri As String, Optional ByVal pbImpTMU As Boolean = False, Optional ByVal pbPagoHonra As Boolean = False) As String
            'WIOR 201312228 AGREGO pbPagoHonra
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim smax1 As Integer
smax1 = 33
    On Error GoTo ErrorImprimeBoleta
    
    vsCadImpresion = ""
    
    'cti3 ferimoro: ERS082-2018
    Dim nCancelado(1) As Variant
    nCancelado(0) = "EXTORNO"
    nCancelado(1) = "FALSE"
    '--------------------------
        
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaProducto(psCtaCod)
    Set oDCred = Nothing
    
    pnSaldoCap = R!nSaldo
    R.Close
    Set R = Nothing
    
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72) ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & ""
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, "", _
            psFecha, psHora, psPersNombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC, , pbPagoHonra, nCancelado)
    
    'Cuerpo de la Boleta
If psCodCMAC = "102" Then
    
Else
    If pbImpTMU = False Then
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(5, "-") & psDescrip & String(7, "-"), 50)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(5, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
        '45 menos la longitud de su etiqueta
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Monto" & Space(smax1 + 7 - Len("Monto") - Len(Format(pnCapPagado, "##,##0.00"))) & Format(pnCapPagado, "##,##0.00"), 51, 0)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Monto" & Space(smax1 + 7 - Len("Monto") - Len(Format(pnCapPagado, "##,##0.00"))) & Format(pnCapPagado, "##,##0.00"), 51, 0) & lsSaltoLin

        'vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion '& lsSaltoLin
        vsCadImpresion = vsCadImpresion & String(40, "-")
        vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Total Pagado " & Space(smax1 + 7 - Len("Total Pagado ") - Len(Format(pnCapPagado, "##,##0.00"))) & Format(pnCapPagado, "##,##0.00"), 51, 0)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Total Pagado " & Space(smax1 + 7 - Len("Total Pagado ") - Len(Format(pnCapPagado, "##,##0.00"))) & Format(pnCapPagado, "##,##0.00"), 40, 0) & lsSaltoLin

        
        'vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        If Not pbPagoHonra Then 'WIOR 20131230
            vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(10) & oImpre.ImpreFormat(pnSaldoCap, 12)
            vsCadImpresion = vsCadImpresion & Space(11) & "Saldo Capital :" & Space(10) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
        End If
        
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Usuario Exto. " & Space(smax1 + 7 - Len("Usuario Exto. ") - Len(psCodUsu)) & psCodUsu, 51, 0)
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Usuario Exto. " & Space(smax1 + 7 - Len("Usuario Exto. ") - Len(psCodUsu)) & psCodUsu, 40, 0) & lsSaltoLin
        
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Usuario Gene. " & Space(smax1 + 7 - Len("Usuario Gene. ") - Len(psCodUserOri)) & psCodUserOri, 51, 0)
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Usuario Gene. " & Space(smax1 + 7 - Len("Usuario Gene. ") - Len(psCodUserOri)) & psCodUserOri, 40, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago " & Space(smax1 + 7 - Len("Prox. Fecha Pago ") - Len(psCodUsu)) & psCodUsu, 51, 0)
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago " & Space(smax1 + 7 - Len("Prox. Fecha Pago ") - Len(psCodUsu)) & psCodUsu, 40, 0) & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & lsSaltoLin
    Else
         vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
        'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
        '45 menos la longitud de su etiqueta
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & String(40, "-")
        'vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        If Not pbPagoHonra Then 'WIOR 20131230
            vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12)
        End If
        'vsCadImpresion = vsCadImpresion & Space(16) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Usuario Exto. :" & Space(7) & oImpre.ImpreFormat(psCodUsu, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Exto. :" & Space(7) & oImpre.ImpreFormat(psCodUsu, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Usuario Gene. :" & Space(7) & oImpre.ImpreFormat(psCodUserOri, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Gene. :" & Space(7) & oImpre.ImpreFormat(psCodUserOri, 12) & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & Chr(27) & "m" & Chr(10)
    End If
End If

    ImprimeBoletaExtorno = vsCadImpresion
    
    Exit Function
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    
End Function

Public Function ImprimeBoletaExtornoAhorros(ByVal psDescrip As String, ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, _
            Optional ByVal pbImpTMU As Boolean = False) As String
            
Dim nFicSal As Integer
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset

    
    vsCadImpresion = ""
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaProducto(psCtaCod)
    Set oDCred = Nothing
    
    pnSaldoCap = R!nSaldo
    R.Close
    Set R = Nothing
    
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72) ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & ""
    'Cabecera
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & lsSaltoLin
 If psCodCMAC = "102" Then
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - AHORROS", 52) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22), 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & lsSaltoLin
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50) & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
 Else
    If pbImpTMU = False Then
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - AHORROS", 52) & Trim(oImpre.ImpreFormat(psNomCmacAbrev & " - AHORROS", 53)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22), 53, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22), 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & oImpre.ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & lsSaltoLin
        
        'Cuerpo de la Boleta
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
        '45 menos la longitud de su etiqueta
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12)
        vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & String(40, "-")
        vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12)
        vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
    Else
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - AHORROS", 52) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22), 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & lsSaltoLin
        
        'Cuerpo de la Boleta
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
        'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
        '45 menos la longitud de su etiqueta
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & String(40, "-")
        'vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
        vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12)
        'vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(pnCapPagado, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsNegritaOff
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & Chr(27) & "m" & Chr(10)
    End If
 End If
    
ImprimeBoletaExtornoAhorros = vsCadImpresion
    
End Function

'Para el Manejo del RFA
Public Function ImprimeBoletaRFA(ByVal pnMovNro As Long, _
                                ByVal psPersNombre As String, _
                                ByVal psNomAge As String, _
                                ByVal pdFecSis As Date, _
                                ByVal psCodUser As String, _
                                ByVal psLpt As String, _
                                ByVal psInstCmac As String, _
                                ByVal psCodCMAC As String) As String

    Dim nCuotas As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim snMonto As Double
    Dim sUser As String
    Dim oDCredito As COMDCredito.DCOMCredDoc
    Dim sImpresion As String
    
    Set oDCredito = New COMDCredito.DCOMCredDoc
    sUser = oDCredito.GetUsuario(pnMovNro)
    Set oDCredito = Nothing
    
    Set oConec = New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "Select Sum(nMonto) as nMonto,cCtaCod"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where nMOvNro=" & pnMovNro & " and cOpeCod not like '107[123456789]%'"
    sSql = sSql & " Group By cCtaCod"
    sSql = sSql & " Order by cCtaCod"
           
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    sImpresion = ""
    Do Until rs.EOF
        sImpresion = sImpresion & ImprimeBoletaExtorno("Extorno RFA", rs!cCtaCod, psPersNombre, _
                   psNomAge, "DOLARES", sGetCuotas(pnMovNro, rs!cCtaCod), pdFecSis, Time, "", IIf(IsNull(rs!nMonto), 0, rs!nMonto), 0#, psCodUser, psLpt, psInstCmac, psCodCMAC, sUser) & lnSaltoPagina
        rs.MoveNext
    Loop
    Set rs = Nothing
                  
    ImprimeBoletaRFA = sImpresion
    
End Function

Function sGetCuotas(ByVal pnMovNro, ByVal psCtaCod As String) As String
   
   Dim oConec As COMConecta.DCOMConecta
   Dim sSql As String
   Dim rs As ADODB.Recordset
   Dim sCuotas As String
   
   
   Set oConec = New COMConecta.DCOMConecta
   oConec.AbreConexion
   sSql = "Select Distinct nNroCuota"
   sSql = sSql & " From MovColDet"
   sSql = sSql & " Where nMovNro=" & pnMovNro & " and cCtaCod='" & psCtaCod & "' and nNroCuota<>0"
   
   Set rs = oConec.CargaRecordSet(sSql)
   oConec.CierraConexion
   Set oConec = Nothing
   
   Do Until rs.EOF
    If Len(sCuotas) = 0 Then
        sCuotas = rs!nNroCuota
    Else
        sCuotas = sCuotas & "," & rs!nNroCuota
    End If
    rs.MoveNext
   Loop
   Set rs = Nothing
   
   sGetCuotas = sCuotas
End Function

Private Sub Class_Initialize()
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
    lsSaltoLin = Chr$(10)
    EspHorxPag = 170
    EspVerxPag = 56
    nPuntPag = 0
    lsTab = Space(1)
    
    Dim oImp As New COMDConstSistema.DCOMImpresoras
    Set oImpresora = New COMFunciones.FCOMVarImpresion
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    
    Set oImpre = New COMFunciones.FCOMImpresion
    Set oCad = New COMFunciones.FCOMCadenas
    bAplicaComision = False 'RIRO20200706
    
End Sub

Private Sub Class_Terminate()
    Set oImpre = Nothing
    Set oCad = Nothing
    Set oImpresora = Nothing
    bAplicaComision = False 'RIRO20200706
End Sub

Private Sub ImprimeCabeceraBoletaPago_NEW(ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, ByVal psNroTransac As String, ByVal psCtaCod As String, _
        ByVal pstipocredito As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, _
        Optional ByVal psDNI As String = "", Optional ByVal bAmpliado As Boolean)
       ' psNomAge = "Sede Institucional"
       'CTI3 ferimoro: ERS082-2018, Optional ByVal bAmpliado As Boolean
        On Error GoTo ErrorImprimeCabeceraBoletaPago
                
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        If psCodCMAC = "102" Then
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Else
            If bAmpliado Then
                vsCadImpresion = ""
                vsCadImpresion = vsCadImpresion & lsNegritaOn
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 51, 0) & Trim(oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 50, 0) & Space(1) & Trim(oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin

                'vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(sDatosBolAmpl(1), 1, 22) & " - " & sDatosBolAmpl(2) & Space(3) & Trim(sDatosBolAmpl(6)), 51, 0) & Trim(oImpre.ImpreFormat(Mid(sDatosBolAmpl(1), 1, 22) & " - " & sDatosBolAmpl(2) & Space(3) & Trim(sDatosBolAmpl(6)), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & lsNegritaOff
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Cliente : " & Mid(psPersNombre, 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Cliente : " & Mid(psPersNombre, 1, 50) & Space(1), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DNI     : " & Mid(psDNI, 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("DNI : " & Mid(psDNI, 1, 50) & Space(1), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Crédito : " & Mid(psCtaCod, 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Crédito : " & Mid(psCtaCod, 1, 50) & Space(1), 53, 0)) & lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 50, 0) & Trim(oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & lsNegritaOff
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50) & Space(5) & "DNI : " & psDNI, 51, 0) & Trim(oImpre.ImpreFormat(Mid(psPersNombre, 1, 50) & Space(5) & "DNI : " & psDNI, 53, 0)) & lsSaltoLin
                vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
            End If
        End If
        
'        vsCadImpresion = vsCadImpresion & lsNegritaOn
'        If psCodCMAC = "102" Then
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & lsNegritaOff
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
'        Else
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 50, 0) & Trim(oImpre.ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(pstipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(oImpre.ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & lsNegritaOff
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(oImpre.ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Mid(psPersNombre, 1, 50) & Space(5) & "DNI : " & psDNI, 51, 0) & Trim(oImpre.ImpreFormat(Mid(psPersNombre, 1, 50) & Space(5) & "DNI : " & psDNI, 53, 0)) & lsSaltoLin
'            vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & oImpre.ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
'        End If
        
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub

'ARCV 28-10-2006
Public Function GeneraArchivoTXT_Educacion(ByVal psPeriodo As String, _
                                            ByVal psTipoPlanilla As String) As String
        On Error GoTo ErrorGeneraArchivoTXT_Educacion
Dim rs As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredDoc

Dim sApeMat As String
Dim sApePat As String
Dim sNom1 As String
Dim sNom2 As String
Dim sApeCas As String
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sCodigoEdu As String
Sql = "Select CE.cCodigoEdu"
Sql = Sql & " from  CredTxtEducacion C"
Sql = Sql & " Inner Join CodigosEduAge CE ON C.cAgeCod = CE.cCodAge"


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set R = oConecta.CargaRecordSet(Sql)

If R.RecordCount > 0 Then
    sCodigoEdu = R!cCodigoEdu
End If

Set oCred = New COMDCredito.DCOMCredDoc

Set rs = oCred.RecuperaDatosColocacConvenio(psTipoPlanilla)

vsCadImpresion = ""

While Not rs.EOF
    Call DevuelveNombrePersonaEnPartes(rs!cPersNombre, sApePat, sApeMat, sApeCas, sNom1, sNom2)
  
    vsCadImpresion = vsCadImpresion & psPeriodo & sCodigoEdu & IIf(rs!cCodModular = "", Space(10), Left(rs!cCodModular & Space(10), 10)) & IIf(rs!cCargo = "", Space(6), Left(rs!cCargo & Space(6), 6)) & IIf(rs!cCARBEN = "", Space(4), Left(rs!cCARBEN & Space(4), 4)) & IIf(rs!cT_Plani = "", Space(1), Left(rs!cT_Plani & Space(1), 1)) & _
                    "151" & String(8 - Len(Replace(CStr(Format(rs!nMontoDes, "#0.00")), ".", "")), "0") & Replace(CStr(Format(rs!nMontoDes, "#0.00")), ".", "") & _
                     sApePat & Space(40 - Len(sApePat)) & sApeMat & Space(40 - Len(sApeMat)) & sNom1 & " " & sNom2 & Space(35 - Len(sNom1 & " " & sNom2)) & Chr(13) & Chr(10)
    rs.MoveNext
Wend
Set oCred = Nothing
        
GeneraArchivoTXT_Educacion = vsCadImpresion
        
Exit Function

ErrorGeneraArchivoTXT_Educacion:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Function


Public Sub DevuelveNombrePersonaEnPartes(ByVal psNomCli As String, _
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
     'psPat = Trim(IIf(psPat = "", "XXXX", psPat))
     'psMat = Trim(IIf(psMat = "", "XXXX", psMat))
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

Public Function RecuperaDatosHojaResumenCreditos(ByVal psCtaCod As String, _
ByRef prsCred As ADODB.Recordset, ByRef prsAge As ADODB.Recordset, _
ByRef pRsRel As ADODB.Recordset, ByRef pcCiaSeg As String, _
ByRef pcNroPol As String, ByRef pnCosRef As Double, _
ByRef pnComDsg As Double, ByRef pnCosCroPag As Double, _
ByRef pnCosHisPag As Double, ByRef pnCosComPag As Double, _
ByRef pnMonInt As Double, ByRef pnPriPol As Double, _
ByRef pnCantAnt As Double, ByRef pnCenCrg As Double, _
ByRef pnSgrDsg1 As Double, ByRef pnSgrDsg2 As Double, _
ByRef pnSegMor As Double, ByRef pnMinCuo As Double, ByRef pnMaxCuo As Double, ByRef pnConScore As Double, ByRef pnCosComPag1 As Double, ByRef psNombreAseguradora As String, _
Optional ByRef pnPriPolInc As Double, Optional ByRef psNombreAseguradoraPol As String = "", Optional ByRef pnPriPolVehic As Double = 0, Optional ByRef bPoliza As Boolean = False) 'MAVM 20100412 Agrego pnCosComPag1
'MAVM 20120215 Agrego pnPriPolInc
'EJVG20131001 psNombreAsegVehic y pnPriPolVehic

    
    Dim oDCred As COMDCredito.DCOMCredito

    Dim ObjCons As New COMDConstantes.DCOMAgencias

    Dim oConst As COMDConstSistema.NCOMConstSistema
    Dim oParam As COMDCredito.DCOMParametro
    Dim oNCred As COMNCredito.NCOMCredito
    Dim MatCalend As Variant
    Dim i As Integer
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim oPol As COMDCredito.DCOMPoliza
    Dim rsAs As ADODB.Recordset
    
    Set ObjCons = New COMDConstantes.DCOMAgencias
    Set prsAge = ObjCons.RecuperaAgencias(Mid(psCtaCod, 4, 2))
    Set ObjCons = Nothing

    Set oDCred = New COMDCredito.DCOMCredito
    Set prsCred = oDCred.RecuperaDatosComunes(psCtaCod)
    Set pRsRel = oDCred.RecuperaRelacPers(psCtaCod)
    Set oDCred = Nothing

    Set oConst = New COMDConstSistema.NCOMConstSistema
    pcCiaSeg = oConst.LeeConstSistema(311)
    Set oConst = Nothing
    Set oParam = New COMDCredito.DCOMParametro
    
    pnCosRef = oParam.RecuperaValorParametro(3070)
    pnComDsg = oParam.RecuperaValorParametro(3071)
    pnCosCroPag = oParam.RecuperaValorParametro(3072)
    pnCosHisPag = oParam.RecuperaValorParametro(3073)
    'pnCosComPag = oParam.RecuperaValorParametro(3074)
    'Comentado Por MAVM 20100412 Memo 841
    pnCosComPag = oParam.RecuperaValorParametro(3122)
    pnCosComPag1 = oParam.RecuperaValorParametro(3123)
    Set rsAs = oParam.RecuperaDatos(102742) 'MADM 20111212
    If rsAs.EOF Then
        psNombreAseguradora = 0
    Else
        psNombreAseguradora = IIf(IsNull(rsAs!cParamDesc), 0, rsAs!cParamDesc)
    End If
     
    Set oParam = Nothing

    Set oNCred = New COMNCredito.NCOMCredito
    'ALPA 20080926****************************************************************
    'MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod)
    
    'pnMonInt = oNCred.MatrizIntCompCalendario(MatCalend)
    pnMonInt = oNCred.ObtenerMontoTotalCalendarioXConcepto(psCtaCod, "1100,1102")
    '*****************************************************************************
    Set oNCred = Nothing
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
     sSql = " Select GP.cNumPoliza" & _
            "    From ColocGarantia CG inner join GarantPoliza GP on CG.cNumGarant=GP.cNumGarant " & _
            "    Where CG.cCtaCod='" & psCtaCod & "' and CG.nEstado<>0 and GP.nEstado<>0 "
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        bPoliza = False
    Else
        bPoliza = True
    End If
    
    sSql = "SELECT nMonto=IsNull(SUM(CD.nMonto),0) FROM ColocCalendDet CD INNER JOIN ColocacCred CC ON CD.cCtaCod= CC.cCtaCod AND CD.nNroCalen=CC.nNroCalen" & _
           " WHERE CD.cCtaCod ='" & psCtaCod & "'" & _
           " AND nColocCalendApl= 1 AND nPrdConceptoCod= " & gColocConceptoCodGastoSeguro7
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnPriPol = 0
    Else
        pnPriPol = IIf(IsNull(rs!nMonto), 0, rs!nMonto)
    End If
    
    sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod=1223"
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnCantAnt = 0
    Else
        pnCantAnt = rs!nValor
    End If
    
    sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod=1221"
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnCenCrg = 0
    Else
        pnCenCrg = rs!nValor
    End If
    
    'MAVM 25112009 ***
    sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod=1234"
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnConScore = 0
    Else
        pnConScore = rs!nValor
    End If
    ' ***
    
    sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod= " & gColocConceptoCodGastoSeguro7
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnSgrDsg1 = 0
        pnSgrDsg2 = 0
    Else
        pnSgrDsg1 = rs!nValor
        pnSgrDsg2 = rs!nValorDosTit
    End If
    
    sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod= 1215"
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnSegMor = 0
        pnMinCuo = 0
        pnMaxCuo = 0
    Else
        pnSegMor = rs!nValor
        pnMinCuo = rs!nMontoMin
        pnMaxCuo = rs!nMontoMax
    End If
    
    'MAVM 20120215 ***
    sSql = "SELECT nMonto=SUM(CD.nMonto) FROM ColocCalendDet CD INNER JOIN ColocacCred CC ON CD.cCtaCod= CC.cCtaCod AND CD.nNroCalen=CC.nNroCalen" & _
           " WHERE CD.cCtaCod ='" & psCtaCod & "'" & _
           " AND nColocCalendApl= 1 AND nPrdConceptoCod in (" & gColocConceptoCodGastoPolizaIncendioHipoteca & "," & gColocConceptoCodGastoPolizaIncendioHipotecaGracia & ")" 'APRI20201020 MEJORA
           '" AND nColocCalendApl= 1 AND nPrdConceptoCod= " & gColocConceptoCodGastoPolizaIncendioHipoteca
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnPriPolInc = 0
    Else
        pnPriPolInc = IIf(IsNull(rs!nMonto), 0, rs!nMonto)
    End If
    '***
    'EJVG20131001 ***
    sSql = "SELECT nMonto=SUM(CD.nMonto) FROM ColocCalendDet CD INNER JOIN ColocacCred CC ON CD.cCtaCod= CC.cCtaCod AND CD.nNroCalen=CC.nNroCalen" & _
           " WHERE CD.cCtaCod ='" & psCtaCod & "'" & _
           " AND nColocCalendApl= 1 AND nPrdConceptoCod= " & gColocConceptoCodGastoPolizaVehiculoMultiriesgo
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.EOF Then
        pnPriPolVehic = 0
    Else
        pnPriPolVehic = IIf(IsNull(rs!nMonto), 0, rs!nMonto)
    End If
    'END EJVG *******
    
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set oPol = New COMDCredito.DCOMPoliza
    'Set Rs = oPol.RecuperaPolizas(0, prsCred!cPersCod)
    Set rs = oPol.RecuperaPolizasCuenta(psCtaCod)
    If Not rs.EOF Then
    'RECO20150213 INC1502040003****************
    Dim x As Integer
    For x = 0 To rs.RecordCount - 1
        If IsNull(rs!cCodPolizaAseg) Then
            pcNroPol = "" & "/"
        Else
            pcNroPol = pcNroPol & rs!cCodPolizaAseg & "/"
        End If
        psNombreAseguradoraPol = IIf(IsNull(rs!cPersNombreAseg), "", rs!cPersNombreAseg) 'EJVG20131001
        rs.MoveNext
    Next
        pcNroPol = Mid(pcNroPol, 1, Len(pcNroPol) - 1)
    Else
        pcNroPol = ""
    End If
    'END***************************************
    Set oPol = Nothing
        
End Function

'*** PEAC 20080412
Public Function RecuperaDatosInformeComercial(ByVal psCtaCod As String, _
ByRef prsCred As ADODB.Recordset)
    
    Dim oDCred As COMDCredito.DCOMCredito
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set prsCred = oDCred.RecuperaDatosInfCome(psCtaCod)
    Set oDCred = Nothing
            
End Function

'*** PEAC 20080412
'Public Function RecuperaDatosBalance(ByVal psNumFte As String, ByRef prsCred As ADODB.Recordset) 'LUCV20160815, Comentó->Según: ERS004-2016
Public Function RecuperaDatosBalance(ByVal psCtaCod As String, ByRef prsCred As ADODB.Recordset) 'LUCV20160815, Comentó->Según: ERS004-2016
    
    Dim oDCred As COMDCredito.DCOMCredito
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set prsCred = oDCred.RecuperaDatosBalGen(psCtaCod)
    Set oDCred = Nothing
            
End Function

'*** PEAC 20080412
Public Function RecuperaDatosInformeComercial01(ByVal psCtaCod As String, _
ByRef prsCred As ADODB.Recordset)
    
    Dim oDCred As COMDCredito.DCOMCredito
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set prsCred = oDCred.RecuperaDatosInfCome01(psCtaCod)
    Set oDCred = Nothing
            
End Function

'**DAOR 20070815
'**Función de que devuelve el reporte de saldos de cartera por analista,
'**obtenido de la base de datos consolidado
Public Function ImprimeResumenSaldosCarteraXAnalistaConsol(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Integer, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pnTipCam As Double, Optional ByVal psTitProductos As String, Optional ByVal pbCondBN As Boolean = False) As String 'DAOR 20070717, se agregí pnTipCam
'ByVal pMatCond As Variant
Dim oDCred As COMDCredito.DCOMCreditos

Dim R As ADODB.Recordset
Dim sCadImp As String
Dim TitDias(5) As String

    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, pMatCond, pMatProd, pnTipCam, pbCondBN)
    Set oDCred = Nothing
    
    '*** PEAC 20080905
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    '*** FIN PEAC
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        TitDias(0) = "Cli   SALDO     " & pnCar1I & "-" & pnCar1F
        TitDias(1) = "Cli   SALDO     " & pnCar2I & "-" & pnCar2F
        TitDias(2) = "Cli   SALDO     " & pnCar3I & "-" & pnCar3F
        TitDias(3) = "Cli   SALDO     > " & pnCar4I

        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        
        '*** PEAC 20080905
        'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
        
        sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 6, 0)
        sCadImp = sCadImp & oImpre.ImpreFormat("CARTERA TOTAL", 20) & "|"
        sCadImp = sCadImp & Space(16) & Space(5) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & Space(5) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & Space(5) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("%MORA", 5) & "|"
        
        '*** PEAC 20080905
        'sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & Space(2) & oImpre.ImpreFormat("%MORA JUD", 9) & "|"
        sCadImp = sCadImp & Space(16) & Space(2) & oImpre.ImpreFormat("%MORA JUD", 9) & "|"
                
        sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(6) & oImpre.ImpreFormat("   Nro   Cli   SALDO  ", 20) & "|"
        sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(0), 21) & "|"
        sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(1), 21) & "|"
        sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(2), 21) & "|"
        sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat(TitDias(3), 21) & "|"
        
        '*** PEAC 20080905
        'sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO     JUD  ", 21) & "|"
        
        'CON DEMANDA
        sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO   C/DMDA.", 21) & "|"
        'SIN DEMANDA
        sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO   S/DMDA.", 21) & "|"
        
        sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO     %MORA", 21) & lnSaltoLinDoc
        
        '*** PEAC 20080905
        'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
        
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                
                '*** PEAC 20080905
                'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & oImpre.ImpreFormat("SUB T:", 6, 0) '
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                '*** PEAC 20080905
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                
                '*** con demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro1, 5, 0) & oImpre.ImpreFormat(R!nNumCli51, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo1, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2)
                '*** sin demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro2, 5, 0) & oImpre.ImpreFormat(R!nNumCli52, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo2, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                
                
                
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & lnSaltoLinDoc
                
                '*** PEAC 20080905
                'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                    
                    sCadImp = sCadImp & oImpre.ImpreFormat("TOTAL:", 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                    'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                    
                    
                    '*** con demanda
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro1, 5, 0) & oImpre.ImpreFormat(R!nNumCli51, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo1, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2)
                    '*** sin demanda
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro2, 5, 0) & oImpre.ImpreFormat(R!nNumCli52, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo2, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                    
                    
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                    
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos   : " & psTitProductos
                    sCadImp = sCadImp & lnSaltoLinDoc
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                    
                    sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat("CARTERA TOTAL", 20) & "|"
                    sCadImp = sCadImp & Space(16) & Space(5) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & Space(5) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(18) & oImpre.ImpreFormat("%MORA JUD", 9) & "|"
                    sCadImp = sCadImp & Space(18) & oImpre.ImpreFormat("%MORA JUD", 9) & "|"
                    
                    sCadImp = sCadImp & Space(16) & Space(6) & oImpre.ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(6) & oImpre.ImpreFormat("   Nro   Cli   SALDO  ", 20) & "|" '
                    sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(0), 21) & "|"
                    sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(1), 21) & "|"
                    sCadImp = sCadImp & "Nro  " & oImpre.ImpreFormat(TitDias(2), 21) & "|"
                    sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat(TitDias(3), 21) & "|"
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO     JUD  ", 21) & "|"
                    'sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO     *    **  ", 25) & "|"
                                          
                    'CON DEMANDA
                    sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO   C/DMDA.", 21) & "|"
                    'SIN DEMANDA
                    sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO   S/DMDA.", 21) & "|"
                     
                    sCadImp = sCadImp & "Nro   " & oImpre.ImpreFormat("Cli   SALDO     %MORA", 21) & lnSaltoLinDoc
                    
                    '*** PEAC 20080905
                    'sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(235, "-") & lnSaltoLinDoc
                End If
            End If
                'madm 2010011 --------------------------------------
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 6, 0)
                If R!nCar6Saldo > 0 Then
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                 
                 '*** con demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro1, 5, 0) & oImpre.ImpreFormat(R!nNumCli51, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo1, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2)
                '*** sin demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro2, 5, 0) & oImpre.ImpreFormat(R!nNumCli52, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo2, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                'IIf(R!nCar6Saldo = 0, 1, R!nCar6Saldo))
                Else
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                 '*** con demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro1, 5, 0) & oImpre.ImpreFormat(R!nNumCli51, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo1, 9, 2) & oImpre.ImpreFormat(R!nCar6Saldo, 4, 2)
                '*** sin demanda
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro2, 5, 0) & oImpre.ImpreFormat(R!nNumCli52, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo2, 9, 2) & oImpre.ImpreFormat(0, 4, 2)
                
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(0, 4, 2)
                End If
                'end madm -----------------------------------------
                
                
                '*** PEAC 20080905
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                
                
                '*** con demanda
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro1, 5, 0) & oImpre.ImpreFormat(R!nNumCli51, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo1, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo1 / R!nCar6Saldo) * 100, 2), 4, 2)
                '*** sin demanda
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro2, 5, 0) & oImpre.ImpreFormat(R!nNumCli52, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo2, 9, 2) & oImpre.ImpreFormat(Round((R!nCar5Saldo2 / R!nCar6Saldo) * 100, 2), 4, 2)
                
                
                'sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
'    '*** PEAC 20080905
'    sCadImp = sCadImp & "*  Mora Créditos Judiciales Con Demanda" & lnSaltoLinDoc
'    sCadImp = sCadImp & "** Mora Créditos Judiciales Sin Demanda" & lnSaltoLinDoc
    
    ImprimeResumenSaldosCarteraXAnalistaConsol = sCadImp
    
End Function

'*** PEAC 20080307
Public Function ImprimeResumenSaldosCarteraXAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal cMatCond As String, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pnTipCam As Double, Optional ByVal psTitProductos As String, Optional ByVal pbCondiBN As Boolean) As String 'DAOR 20070717, se agregí pnTipCam....GITU se agrego pbCondiBN 05/04/2008
    'ByVal pMatCond As Variant
Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim sCadImp As String
  
    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalista(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, cMatCond, pMatProd, pnTipCam, pbCondiBN)
    Set oDCred = Nothing
    
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
        
        '**DAOR 20070427, Titulos de Columnas del Reporte
        sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 6, 0)
        sCadImp = sCadImp & oImpre.ImpreFormat("CARTERA TOTAL", 20) & "|"
        sCadImp = sCadImp & Space(20) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(21) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(21) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(21) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(21) & oImpre.ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(21) & oImpre.ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(6) & oImpre.ImpreFormat("   Nro   Cli   SALDO  ", 20) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     1 - 7", 25) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     8 -15", 26) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     16-30", 26) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     > 30 ", 26) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     JUD  ", 26) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("Nro  Cli  SALDO     %MORA", 26) & lnSaltoLinDoc
        
        sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & oImpre.ImpreFormat("SUB T:", 6, 0) '
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                '*******************************************************
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & oImpre.ImpreFormat("TOTAL:", 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                    '*********************************************************
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    '**DAOR 20070427, Titulos de Columnas del Reporte
                    sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos   : " & psTitProductos
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat("CARTERA TOTAL", 20) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & oImpre.ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(6) & oImpre.ImpreFormat("   Nro   Cli   SALDO  ", 20) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     1 - 7", 21) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     8 -15", 21) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     16-30", 21) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     > 30 ", 21) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     JUD  ", 21) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("Nro   SALDO     %MORA", 21) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(199, "-") & lnSaltoLinDoc
                End If
            End If
                '**DAOR 20070512
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 4, 0) & Space(2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar6Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli6, 5, 0) & oImpre.ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli1, 5, 0) & oImpre.ImpreFormat(R!nCar1Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar2Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli2, 5, 0) & oImpre.ImpreFormat(R!nCar2Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar3Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli3, 5, 0) & oImpre.ImpreFormat(R!nCar3Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar4Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli4, 5, 0) & oImpre.ImpreFormat(R!nCar4Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar5Nro, 5, 0) & oImpre.ImpreFormat(R!nNumCli5, 5, 0) & oImpre.ImpreFormat(R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & oImpre.ImpreFormat(R!nNumCli1 + R!nNumCli2 + R!nNumCli3 + R!nNumCli4 + R!nNumCli5, 6, 0) & oImpre.ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & oImpre.ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXAnalista = sCadImp
    
End Function


'**DAOR 20070815
'**Función de que devuelve el reporte consolidado de cartera por analista,
'**obtenido de la base de datos consolidado
Public Function ImprimeConsolidadoCarteraXAnalistaConsol(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal pMatProd As Variant, _
    Optional ByVal psNomCmac As String = "", Optional ByVal pnTipCam As Double, Optional ByVal psTitProductos As String) As String
Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim sCadImp As String

    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONSOLIDADO DE CARTERA POR ANALISTA", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
        
        sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 7, 0)
        sCadImp = sCadImp & oImpre.ImpreFormat("  SALDO CARTERA", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  NUEVOS", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  RECURRENTES", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  PARALELOS", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  REFINANCIADOS", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  AMPLIADO", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  AUTOMATICOS", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("  ADICIONALES", 17) & "|" & lnSaltoLinDoc 'JUEZ 20130603
        sCadImp = sCadImp & Space(7)
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|" & lnSaltoLinDoc 'JUEZ 20130603
        sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
                sCadImp = sCadImp & oImpre.ImpreFormat("SUB T:", 6, 0) '
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRecurrente, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoParalelo, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRefinanciado, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAmpliado, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAutomatico, 10, 2)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAdicional, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAdicional, 10, 2) 'JUEZ 20130603
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
                    sCadImp = sCadImp & oImpre.ImpreFormat("TOTAL:", 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico + R!nCantAdicional, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico + R!nMontoAdicional, 10, 2) 'JUEZ 20130603
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRecurrente, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoParalelo, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRefinanciado, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAmpliado, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAutomatico, 10, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAdicional, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAdicional, 10, 2) 'JUEZ 20130603
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos   : " & psTitProductos
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & oImpre.ImpreFormat("ANAL", 7, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat("  SALDO CARTERA", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  NUEVOS", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  RECURRENTES", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  PARALELOS", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  REFINANCIADOS", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  AMPLIADO", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  AUTOMATICOS", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("  ADICIONAL", 17) & "|" & lnSaltoLinDoc 'JUEZ 20130603
                    sCadImp = sCadImp & Space(7)
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & oImpre.ImpreFormat("NUM.   MONTO     ", 17) & "|" & lnSaltoLinDoc 'JUEZ 20130603
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc 'JUEZ 20130603
                End If
            End If
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 6, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico + R!nCantAdicional, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico + R!nMontoAdicional, 10, 2) 'JUEZ 20130603
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoNuevo, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRecurrente, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoParalelo, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoRefinanciado, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAmpliado, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAutomatico, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCantAdicional, 6, 0) & Space(1) & oImpre.ImpreFormat(R!nMontoAdicional, 10, 2) 'JUEZ 20130603
            sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeConsolidadoCarteraXAnalistaConsol = sCadImp
    
End Function

'*** PEAC 20070822
Public Function ImprimeActaComiteCredApro(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatProd As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc 'DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vProd As String

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set R = oDCred.RecuperaActaComiteCredApro(psCodAge, pdFecIni, pdFecFin, pMatProd)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "ACTAS DE COMITE DE CREDITOS APROBADOS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        'sCadImp = sCadImp & "Agencia     : " & R!cAgeDescripcion & lnSaltoLinDoc
        'sCadImp = sCadImp & "Agencia     : " & psNomAge & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(146, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                     FECHA       MONTO" & lnSaltoLinDoc
        sCadImp = sCadImp & "  CREDITO              CLIENTE                       APROBAC.    APROBADO       CUOTAS PLAZO             LINEA                           ANALISTA" & lnSaltoLinDoc
        sCadImp = sCadImp & String(146, "-") & lnSaltoLinDoc
    
    Do While Not R.EOF
        sCadImp = sCadImp & "Producto  : " & R!Producto
        sCadImp = sCadImp & lnSaltoLinDoc
        vProd = R!Producto
        Do While R!Producto = vProd
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 30)
            sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!dPrdEstado, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nMonto, 12, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nCuotas, 3, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cConsDescripcion, 12)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Plazo, 3, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Descri_Linea, 34)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 4)
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
        Loop
    Loop
    
    R.Close
    Set R = Nothing
    ImprimeActaComiteCredApro = sCadImp
    
End Function

'*** PEAC 20070822
Public Function ImprimeRepXTipoCondProd(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psNomAge As String, _
ByVal pMatProd As Variant, ByVal psMoneda As String, Optional ByVal psNomCmac As String = "", _
Optional ByVal pnTipCam As Double, Optional ByVal pMatAnalistas As Variant) As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vCondi As String
Dim lnTotMonSol As Double, J As Integer

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaRepXTipoCondProd(psCodAge, pdFecIni, pdFecFin, pMatProd, pnTipCam, psMoneda, pMatAnalistas)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE POR TIPO DE CONDICION DE PRODUCTO", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & "Tipo Cambio  : " & pnTipCam & lnSaltoLinDoc
        'sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                                            MONTO            MONTO      FECHA                " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CREDITO            CLIENTE                                 MONEDA     DESEMBOLSADO     SOLES      VIGENCIA    ANALISTA " & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
    
    Do While Not R.EOF
    
        sCadImp = sCadImp & "Condición Producto : " & R!cCondiProd & lnSaltoLinDoc
        vCondi = R!cCondiProd
        lnTotMonSol = 0
        J = 0
        Do While R!cCondiProd = vCondi
            J = J + 1
            sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cTitular, 40)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cMoneda, 6)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoCol, 12, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoColSol, 12, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cAnalista, 4)
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotMonSol = lnTotMonSol + R!nMontoColSol
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
        Loop
        sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(89)
        sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMonSol, 12, 2)
        sCadImp = sCadImp & lnSaltoLinDoc
    Loop
    
    R.Close
    Set R = Nothing
    ImprimeRepXTipoCondProd = sCadImp
       
End Function

'*** PEAC 20070822

Public Function ImprimeRepDetalleCuotasXCobrar(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal psMoneda As String, _
Optional ByVal psNomCmac As String = "", Optional ByVal pMatAnalistas As Variant) As String

Dim oDCred As COMDCredito.DCOMCredDoc 'DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vMoney As String
Dim vAna As String
Dim J As Integer
Dim lnTotSaldo As Double

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set R = oDCred.RecuperaRepDetalleCuotasXCobrar(psCodAge, pdFecIni, pdFecFin, pdFecSis, pMatProd, psMoneda, pMatAnalistas)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE DETALLE DE CUOTAS POR COBRAR", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        'sCadImp = sCadImp & "Agencia     : " & R!cAgeDescripcion & lnSaltoLinDoc
        'sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                                     MONTO               SALDO CAP.         FECHA" & lnSaltoLinDoc
        sCadImp = sCadImp & "     ITEM  CREDITO              CLIENTE                              DESEMBOL.  CUOTA    CUOTA      ATRASO  VMTO." & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
    
    Do While Not R.EOF
        sCadImp = sCadImp & "Analista : " & R!ANALISTA & lnSaltoLinDoc
        vAna = R!ANALISTA
            
            Do While R!ANALISTA = vAna
                sCadImp = sCadImp & "       Moneda   : " & R!Moneda & lnSaltoLinDoc
                vMoney = R!Moneda
                J = 0
                lnTotSaldo = 0
                Do While R!Moneda = vMoney
                    J = J + 1
                    sCadImp = sCadImp & Space(5)
                    sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!Credito, 18)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 30)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!Desembolso, 12, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!Cuota, 5, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!SaldoCapital, 12, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(R!Atraso, 6, 0)
                    sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!fechavenc, "dd/mm/yyyy"), 10)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    lnTotSaldo = lnTotSaldo + R!SaldoCapital
                    R.MoveNext
                    If R.EOF Then
                       Exit Do
                    End If
                Loop
                sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & Space(81)
                sCadImp = sCadImp & oImpre.ImpreFormat(lnTotSaldo, 12, 2)
                sCadImp = sCadImp & lnSaltoLinDoc
                If R.EOF Then
                    Exit Do
                End If
            Loop
    Loop
    
    R.Close
    Set R = Nothing
    ImprimeRepDetalleCuotasXCobrar = sCadImp
    
End Function

'   PEAC 20070822  -- opcion 108124
'   MADM 20091711  -- opcion 108124 - MADM 20100722 cMatCondicion
Public Function ImprimeRepCreditosXCampanha(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal pMatProd As Variant, _
ByVal psMoneda As String, Optional ByVal psNomCmac As String = "", _
Optional ByVal pMatAnalistas As Variant, Optional ByVal pMatCampanhas As Variant, Optional ByVal nTipo As Integer = 1, _
Optional ByVal cMatCondicion As String) As String

'CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias,
'gsCodAge , gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda,
'gsNomCmac, matAnalista, matCampanha

'ByVal pMatAgencias As Variant


Dim oDCred As COMDCredito.DCOMCredDoc 'DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vMoney As String
Dim vCampa As String
Dim vage As String
Dim vSituaC As String
Dim J As Integer
Dim lnTotMonto As Double, lnTotCapPa As Double, lnTotSaldo As Double

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set R = oDCred.RecuperaRepCreditosXCampanha(psCodAge, pMatAgencias, pdFecIni, pdFecFin, psMoneda, pMatProd, pMatAnalistas, pMatCampanhas, nTipo, cMatCondicion)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    
    sCadImp = ""

        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE CREDITOS POR CAMPAÑAS Y/O PRODUCTOS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(180, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "                                                            MONTO              CAPITAL      SALDO      FECHA        DIAS               NOMBRE" & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CREDITO               CLIENTE                         DESEMBOLSO         PAGADO       CAPITAL    VIGENCIA     ATRASO ANALISTA    INSTITUCION" & lnSaltoLinDoc
        sCadImp = sCadImp & String(180, "-") & lnSaltoLinDoc
        vMoney = ""
        vSituaC = ""
        vCampa = ""
        vage = ""
        Dim nChangeAgencia As Integer
        Dim nChangeMoneda As Integer
        Dim nChangeCampanha As Integer
        nChangeAgencia = 0
        nChangeMoneda = 0
        nChangeCampanha = 0
        J = 0
        lnTotMonto = 0
        lnTotCapPa = 0
        lnTotSaldo = 0
        
    Do While Not R.EOF
        If R!Agencia <> vage Then
            nChangeAgencia = 1
        End If
        
        If vMoney <> R!Moneda Then
            nChangeMoneda = 1
        End If
        
        If R!Campanha <> vCampa Or nChangeMoneda = 1 Then
            nChangeCampanha = 1
        End If
        
        If R!Situacion <> vSituaC Or nChangeCampanha = 1 Then
          
                If vSituaC <> "" Then
                    sCadImp = sCadImp & String(180, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(56)
                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMonto, 12, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotCapPa, 12, 2)
                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotSaldo, 12, 2)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    'sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                End If
         If nChangeAgencia = 1 Then
            sCadImp = sCadImp & "AGENCIA  : " & R!Agencia & lnSaltoLinDoc
         End If
         
         If nChangeMoneda = 1 Then
            sCadImp = sCadImp & "MONEDA  : " & R!Moneda & lnSaltoLinDoc
         End If
         
         If nChangeCampanha = 1 Then
            sCadImp = sCadImp & "  CAMPAÑA  : " & R!Campanha & lnSaltoLinDoc
         End If
         sCadImp = sCadImp & "    SITUACIÓN : " & R!Situacion & lnSaltoLinDoc
         sCadImp = sCadImp & String(180, "-") & lnSaltoLinDoc
         lnTotMonto = 0
         lnTotCapPa = 0
         lnTotSaldo = 0
         J = 0
        End If
        nChangeMoneda = 0
        nChangeCampanha = 0
        J = J + 1
        sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 30)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoCol, 12, 2)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!CapitalPagado, 12, 2)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!nSaldo, 12, 2)
        sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 10)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!Atraso, 6, 0)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 8)
        sCadImp = sCadImp & oImpre.ImpreFormat(R!Convenio, 20)
        sCadImp = sCadImp & lnSaltoLinDoc
        
        lnTotMonto = lnTotMonto + R!nMontoCol
        lnTotCapPa = lnTotCapPa + R!CapitalPagado
        lnTotSaldo = lnTotSaldo + R!nSaldo
        
        vMoney = R!Moneda
        vSituaC = R!Situacion
        vCampa = R!Campanha
        vage = R!Agencia
        R.MoveNext
        If R.EOF Then
            Exit Do
        End If
                                             
'                Do While R!Moneda = vMoney
'                    sCadImp = sCadImp & "      Campaña : " & R!Campanha & lnSaltoLinDoc
'                    vCampa = R!Campanha
'                    vSituaC = R!Situacion
'                    J = 0
'                    lnTotMonto = 0
'                    lnTotCapPa = 0
'                    lnTotSaldo = 0
'
'                            Do While R!Campanha = vCampa  'madm situacion
'                             sCadImp = sCadImp & "      Situacion : " & R!Situacion & lnSaltoLinDoc
'
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 30)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoCol, 12, 2)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!CapitalPagado, 12, 2)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!nSaldo, 12, 2)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 10)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!Atraso, 6, 0)
'                                             sCadImp = sCadImp & oImpre.ImpreFormat(R!Analista, 4)
'                                             sCadImp = sCadImp & lnSaltoLinDoc
'
'
'                                             lnTotMonto = lnTotMonto + R!nMontoCol
'                                             lnTotCapPa = lnTotCapPa + R!CapitalPagado
'                                             lnTotSaldo = lnTotSaldo + R!nSaldo
'                                             R.MoveNext
'                                             If R.EOF Then
'                                                 Exit Do
'                                             End If
'                               Loop
'
'                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & Space(56)
'                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMonto, 12, 2)
'                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotCapPa, 12, 2)
'                    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotSaldo, 12, 2)
'                    sCadImp = sCadImp & lnSaltoLinDoc
'
'                    If R.EOF Then
'                        Exit Do
'                    End If
'              Loop
        
         Loop
        If (lnTotMonto + lnTotCapPa + lnTotSaldo) > 0 Then
            sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & Space(56)
            sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMonto, 12, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(lnTotCapPa, 12, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(lnTotSaldo, 12, 2)
            sCadImp = sCadImp & lnSaltoLinDoc
        End If
    
    R.Close
    Set R = Nothing
    ImprimeRepCreditosXCampanha = sCadImp

End Function

'   PEAC 20070921  -- opcion 108209

Public Function ImprimeRepConvCasillero(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal pMatProd As Variant, _
ByVal psMoneda As String, ByVal pMatInstitucion As Variant, Optional ByVal psNomCmac As String = "", _
Optional ByVal pMatAnalistas As Variant, Optional ByVal pMatCampanhas As Variant, _
Optional ByVal pMatGastos As Variant) As String

Dim oDCred As COMDCredito.DCOMCredDoc 'DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vMoney As String
Dim vInsti As String
Dim J As Integer
Dim lnTotMonto As Double, lnTotMontoTot As Double, lnTotSaldo As Double

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set R = oDCred.RecuperaRepConvCasillero(psCodAge, pMatAgencias, pdFecIni, pdFecFin, psMoneda, pMatProd, pMatAnalistas, pMatCampanhas, pMatInstitucion, pMatGastos)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE COBRANZA DIARIA DE CONVENIOS (CASILLERO)", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
       sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                         NUMERO      MONTO       FECHA               DESCRIPCION" & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CREDITO                 CLIENTE                    CUOTA       PAGADO      PAGO    ANALISTA    GASTOS     " & lnSaltoLinDoc
        sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    
    lnTotMontoTot = 0
    Do While Not R.EOF
        sCadImp = sCadImp & "INSTITUCION : " & R!Institucion & lnSaltoLinDoc
        vInsti = R!Institucion
    
            J = 0
            lnTotMonto = 0
            'lnLineas = 0
            Do While R!Institucion = vInsti
                J = J + 1
                sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cPersNombre, 30)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nCuota, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoPagado, 12, 2, True)
                sCadImp = sCadImp & oImpre.ImpreFormat(Format(R!dPago, "dd/mm/yyyy"), 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 4)
                sCadImp = sCadImp & Space(4)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!CDescripcion, 30)
                sCadImp = sCadImp & lnSaltoLinDoc
                lnTotMonto = lnTotMonto + R!nMontoPagado
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
            Loop
            sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & Space(15) & "TOTAL POR INSTITUCION..."
            sCadImp = sCadImp & Space(21)
            sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMonto, 12, 2, True)
            sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
            lnTotMontoTot = lnTotMontoTot + lnTotMonto
            If R.EOF Then
                Exit Do
            End If
    Loop
    sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(15) & "TOTAL GENERAL..."
    sCadImp = sCadImp & Space(29)
    sCadImp = sCadImp & oImpre.ImpreFormat(lnTotMontoTot, 12, 2, True)
    sCadImp = sCadImp & lnSaltoLinDoc
    
    R.Close
    Set R = Nothing
    ImprimeRepConvCasillero = sCadImp

End Function

'peac 20071228 OPERACION 128329

Public Function ImprimeComVigEEFF(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc 'DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String


    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosListadoComVigEEFF(pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If

    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE CREDITOS COMERCIALES VIGENTES", psNomCmac, , True, , "CON ULTIMA FECHA ESTADOS FINANCIEROS")
        sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "      NOMBRE DEL                                            NUMERO CREDITO      CONDICION DEL                 ULT. FECHA       " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CLIENTE                                               VIGENTE             CREDITO              ANALISTA EE.FF      MONEDA" & lnSaltoLinDoc
        sCadImp = sCadImp & String(126, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    
    Do While Not R.EOF
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        
        J = 0
        Do While R!CodAge = vage
                J = J + 1
                sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cPersNombre, 50)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Condicion, 20)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 4)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!ult_fec_eeff) <= "1920", "  /  /    ", Format(R!ult_fec_eeff, "dd/mm/yyyy")), 12)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Moneda, 12, 0)
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
        Loop

    Loop
      
    R.Close
    Set R = Nothing
    ImprimeComVigEEFF = sCadImp

End Function

'*** PEAC 20080215 OPERACION 108141

Public Function ImprimeRepDiasTranscDesdeSoli(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    Screen.MousePointer = 11

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosDiasTranscDesdeSoli(pdFecIni, pdFecFin, pMatAgencias, pnDiasIni, pnDiasFin)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If

    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "DIAS TRANSCURRIDOS DESDE LA SOLICITUD", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & "PROMEDIO DE DIAS TRANSCURRIDOS: " & Trim(Str(R!Promedio)) & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO   DIF.                                                                                     " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIAS    DIRECCION                TELEFONO      ACTIVIDAD                       ANLT. DIAS" & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 9
    Do While Not R.EOF
        i = i + 1
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        J = 0
        Do While R!CodAge = vage
                i = i + 1
                J = J + 1
                sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 35)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Soli) <= "1920", "  /  /    ", Format(R!Fecha_Soli, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Descri_Est_Act, 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Est_Act) <= "1920", "  /  /    ", Format(R!Fecha_Est_Act, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & Space(4) & oImpre.ImpreFormat(R!DiasDif, 3, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Direccion, 27)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Telefono, 12)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Actividad, 30)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 5)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Dias, 3, 0)
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                                                                                
                If R.EOF Then
                    Exit Do
                End If
                                
                If i = 56 Then
                    sCadImp = sCadImp & lnSaltoPagina
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "DIAS TRANSCURRIDOS DESDE LA SOLICITUD", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
                    sCadImp = sCadImp & "PROMEDIO DE DIAS TRANSCURRIDOS: " & Trim(Str(R!Promedio)) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO   DIF.                                                                                     " & lnSaltoLinDoc
                    sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIAS    DIRECCION                TELEFONO      ACTIVIDAD                       ANLT. DIAS" & lnSaltoLinDoc
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                    i = 9
                End If
        Loop
    Loop

    R.Close
    Set R = Nothing
    ImprimeRepDiasTranscDesdeSoli = sCadImp
    
    Screen.MousePointer = 0
    
End Function

'*** PEAC 20080215 OPERACION 108142

Public Function ImprimeRepSoliRechazadas(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    Screen.MousePointer = 11

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosSoliRechazadas(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If

    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES RECHAZADAS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO   DIF.                                                                              " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIAS  DIRECCION                TELEFONO      ACTIVIDAD                       ANLT." & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 8
    Do While Not R.EOF
        i = i + 1
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        J = 0
        Do While R!CodAge = vage
                i = i + 1
                J = J + 1
                sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 35)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Soli) <= "1920", "  /  /    ", Format(R!Fecha_Soli, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Descri_Est_Act, 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Est_Act) <= "1920", "  /  /    ", Format(R!Fecha_Est_Act, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & Space(4) & oImpre.ImpreFormat(R!DiasDif, 3, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Direccion, 27)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Telefono, 12)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Actividad, 30)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 5)
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                                                                                
                If R.EOF Then
                    Exit Do
                End If
                                
                If i = 56 Then
                    sCadImp = sCadImp & lnSaltoPagina
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES RECHAZADAS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO   DIF.                                                                              " & lnSaltoLinDoc
                    sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIAS  DIRECCION                TELEFONO      ACTIVIDAD                       ANLT." & lnSaltoLinDoc
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                    i = 8
                End If
                                
        Loop
        
    Loop

    R.Close
    Set R = Nothing
    ImprimeRepSoliRechazadas = sCadImp
    Screen.MousePointer = 0
End Function

'*** PEAC 20080215 OPERACION 108143

Public Function ImprimeRepSoloEstadoSoli(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, _
ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    Screen.MousePointer = 11

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosSoloEstadoSoli(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If

    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES PENDIENTES", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO                                                                               " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIRECCION                TELEFONO      ACTIVIDAD                       ANLT." & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 8
    Do While Not R.EOF
        i = i + 1
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        J = 0
        Do While R!CodAge = vage
                i = i + 1
                J = J + 1
                sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 35)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Soli) <= "1920", "  /  /    ", Format(R!Fecha_Soli, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Descri_Est_Act, 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Est_Act) <= "1920", "  /  /    ", Format(R!Fecha_Est_Act, "dd/mm/yyyy")), 10)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Direccion, 27)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Telefono, 12)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!Actividad, 30)
                sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 5)
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                                                                                
                If R.EOF Then
                    Exit Do
                End If
                                
                If i = 56 Then
                    sCadImp = sCadImp & lnSaltoPagina
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES PENDIENTES", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & "                                            FECHA       ESTADO      FECHA ESTADO                                                                               " & lnSaltoLinDoc
                    sCadImp = sCadImp & "ITEM  CLIENTE                               SOLICITUD   ACTUAL      ACTUAL         DIRECCION                TELEFONO      ACTIVIDAD                       ANLT." & lnSaltoLinDoc
                    sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                    i = 8
                End If
                                
        Loop
       
    Loop

    R.Close
    Set R = Nothing
    ImprimeRepSoloEstadoSoli = sCadImp
    Screen.MousePointer = 0
End Function

'*** PEAC 20080215 OPERACION 108144
Public Function ImprimeRepSoliProcesadas(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosSoliProcesadas(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES PROCESADAS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                                        FECHA       FECHA        DIF.  MONTO          MONTO     ESTADO                                   " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CUENTA               MONEDA  CLIENTE                              SOLICITUD   APROBACION   DIAS  SOLICITADO     APROBADO  ACTUAL          CAMPAÑA          ANALISTA" & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 8
    Do While Not R.EOF
        i = i + 1
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        J = 0
        Do While R!CodAge = vage
            i = i + 1
            J = J + 1
            sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Cuenta, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Moneda, 6)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Cliente, 35)
            sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Soli) <= "1920", "  /  /    ", Format(R!Fecha_Soli, "dd/mm/yyyy")), 10)
            sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!Fecha_Apro) <= "1920", "  /  /    ", Format(R!Fecha_Apro, "dd/mm/yyyy")), 10)
            sCadImp = sCadImp & Space(4) & oImpre.ImpreFormat(R!DiasDif, 3, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!montosoli, 10, 2, True)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!MontoApro, 10, 2, True)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Descri_Est_Act, 15)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Campanha, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!ANALISTA, 6)
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext

            If R.EOF Then
                Exit Do
            End If

            If i = 56 Then
                sCadImp = sCadImp & lnSaltoPagina
                Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SOLICITUDES PROCESADAS", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
                sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & "                                                                        FECHA       FECHA        DIF.  MONTO          MONTO     ESTADO                                   " & lnSaltoLinDoc
                sCadImp = sCadImp & "ITEM  CUENTA               MONEDA  CLIENTE                              SOLICITUD   APROBACION   DIAS  SOLICITADO     APROBADO  ACTUAL          CAMPAÑA          ANALISTA" & lnSaltoLinDoc
                sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                i = 8
            End If
        Loop
    Loop

    R.Close
    Set R = Nothing
    ImprimeRepSoliProcesadas = sCadImp
    Screen.MousePointer = 0
End Function

'*** PEAC 20080221 OPERACION 108336
Public Function ImprimeRepCliConDistinTiposCred(ByVal pMatAgencias As Variant, _
ByVal pMatProd As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    If UBound(pMatProd) = 0 Then
        MsgBox "Seleccione por lo menos un producto", vbInformation, "Atención"
        Exit Function
    End If

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosCliConDistinTiposCred(pMatAgencias, pMatProd)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CLIENTES CON DISTINTOS TIPOS DE CREDITOS", psNomCmac, , True, , " AL " & Format(pdFecSis, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                                                                           MONTO            SALDO     FECHA                        " & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  NOMBRE CLIENTE                             No CREDITO          MONEDA  TIPO CREDITO                  DESEMBOLSADO     CAPITAL   DESEMBOLSO  AGENCIA  ANALISTA" & lnSaltoLinDoc
        sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 8
    J = 0
    Do While Not R.EOF
            i = i + 1
            J = J + 1
            sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cPersNombre, 40)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Moneda, 6)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!tipoprod, 30)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nMontoCol, 10, 2, True)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nSaldo, 10, 2, True)
            sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!dVigencia) <= "1950", "  /  /    ", Format(R!dVigencia, "dd/mm/yyyy")), 10)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!Agencia, 6)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 6)
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext

            If R.EOF Then
                Exit Do
            End If

            If i = 56 Then
                sCadImp = sCadImp & lnSaltoPagina
                Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CLIENTES CON DISTINTOS TIPOS DE CREDITOS", psNomCmac, , True, , " AL " & Format(pdFecSis, "dd/mm/yyyy"))
                sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & "                                                                                                           MONTO            SALDO     FECHA                        " & lnSaltoLinDoc
                sCadImp = sCadImp & "ITEM  NOMBRE CLIENTE                             No CREDITO          MONEDA  TIPO CREDITO                  DESEMBOLSADO     CAPITAL   DESEMBOLSO  AGENCIA  ANALISTA" & lnSaltoLinDoc
                sCadImp = sCadImp & String(163, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                i = 8
            End If
    Loop

    R.Close
    Set R = Nothing
    ImprimeRepCliConDistinTiposCred = sCadImp
    Screen.MousePointer = 0
End Function

'*** PEAC 20080515
Public Function ImprimeCodHojEval() As String

    Dim rsCodHojEval As ADODB.Recordset
    Dim lnTotPatri As Double
    
    Dim oDAgencia As COMDConstantes.DCOMAgencias
    
    Dim oNCredito As COMNCredito.NCOMCredDoc
    Dim oDCredito As COMDCredito.DCOMCredDoc
    Dim oDPer As New COMDPersona.DCOMPersonas

    Dim sCadImp As String, lcTitulo As String
    
    Dim i As Integer, J As Integer

    Screen.MousePointer = 11
       
    Set oDCredito = New COMDCredito.DCOMCredDoc
        Set rsCodHojEval = oDCredito.ObtenerCodHojEval()
    Set oDCredito = Nothing
            
    If rsCodHojEval.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    lnTotPatri = 0
    
    sCadImp = ""
        sCadImp = sCadImp & "CAJA MAYNAS" & lnSaltoLinDoc
        sCadImp = sCadImp & "RUC : 20103845328" & lnSaltoLinDoc & lnSaltoLinDoc
        
        lcTitulo = "CODIGOS DE LA HOJA DE EVALUACION"
        sCadImp = sCadImp & Space((EspHorxPag - Len(lcTitulo)) / 2 - 18) & lcTitulo & lnSaltoLinDoc
        sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc & lnSaltoLinDoc
        
    sCadImp = sCadImp & "  CODIGO         DESCRIPCION" & lnSaltoLinDoc
    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
    Do While Not rsCodHojEval.EOF
        sCadImp = sCadImp & oImpre.ImpreFormat(rsCodHojEval!cCodHojEval, 15)
        sCadImp = sCadImp & oImpre.ImpreFormat(rsCodHojEval!CDescripcion, 50)
        sCadImp = sCadImp & lnSaltoLinDoc
        rsCodHojEval.MoveNext
        If rsCodHojEval.EOF Then
            Exit Do
        End If
    Loop
    rsCodHojEval.Close
    Set rsCodHojEval = Nothing
    
    ImprimeCodHojEval = sCadImp
    Screen.MousePointer = 0
End Function

Public Function ImprimeDDJJPatrimonial(ByVal pcCuenta As String, ByVal pdFecSis As Date, ByVal pcCodAge As String) As String

'    Dim rsInmue As ADODB.Recordset
'    Dim rsOtros As ADODB.Recordset
'    Dim rsVehic As ADODB.Recordset
'    Dim rsDatosDDJJ As ADODB.Recordset
'    Dim rsAgencia As ADODB.Recordset
'    Dim lnTotPatri As Double
'
'    Dim oDAgencia As COMDConstantes.DCOMAgencias
'
'    Dim oNCredito As COMNCredito.NCOMCredDoc
'    Dim oDCredito As COMDCredito.DCOMCredDoc
'    Dim oDPer As New COMDPersona.DCOMPersonas
'    Dim oDCred As COMDCredito.DCOMCredito
'
'    Dim sCadImp As String, lcTitulo As String
'
'    Dim i As Integer, J As Integer
'
'    Screen.MousePointer = 11
'
'    Set oDAgencia = New COMDConstantes.DCOMAgencias
'    Set rsAgencia = oDAgencia.RecuperaAgencias(pcCodAge)
'    Set oDAgencia = Nothing
'
'    Set oDCredito = New COMDCredito.DCOMCredDoc
'        Set rsDatosDDJJ = oDCredito.ObtenerDatosDDJJPatri(pcCuenta)
'    Set oDCredito = Nothing
'
'    If rsDatosDDJJ.RecordCount = 0 Then
'        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
'        Screen.MousePointer = 0
'        Exit Function
'    End If
'
'     Set oDCredito = New COMDCredito.DCOMCredDoc
'        Set rsInmue = oDCredito.ObtenerPatrimonio(rsDatosDDJJ!cPersCod, "1")
'        Set rsVehic = oDCredito.ObtenerPatrimonio(rsDatosDDJJ!cPersCod, "2")
'        Set rsOtros = oDCredito.ObtenerPatrimonio(rsDatosDDJJ!cPersCod, "3")
'    Set oDCredito = Nothing
'
'    lnTotPatri = 0
'
'    sCadImp = ""
'        'Call ImprimeCabeceraDocumento(sCadImp, "", pdFecSis, "", "DECLARACION JURADA PATRIMONIAL", "", , True, , "")
'        sCadImp = sCadImp & "CAJA MAYNAS" & lnSaltoLinDoc
'        sCadImp = sCadImp & "RUC : 20103845328" & lnSaltoLinDoc & lnSaltoLinDoc
'
'        lcTitulo = "DECLARACION JURADA PATRIMONIAL"
'        sCadImp = sCadImp & Space((EspHorxPag - Len(lcTitulo)) / 2 - 18) & lcTitulo & lnSaltoLinDoc
'        sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc & lnSaltoLinDoc
'        If Len(rsDatosDDJJ!DNI) > 0 Then
'            sCadImp = sCadImp & "YO, " & Trim(rsDatosDDJJ!cPersNombre) & ", IDENTIFICADO CON DNI No " & Trim(rsDatosDDJJ!DNI) & " DE ESTADO CIVIL, " & IIf(Len(rsDatosDDJJ!EstadoCiv) > 0, Trim(rsDatosDDJJ!EstadoCiv), "**********") & lnSaltoLinDoc
'            sCadImp = sCadImp & "Y (nombre del conyugue, si fuese el caso) " & IIf(Len(rsDatosDDJJ!Conyuge) > 0, Trim(rsDatosDDJJ!Conyuge), IIf(Len(rsDatosDDJJ!Codeudor) > 0, Trim(rsDatosDDJJ!Codeudor), "**********")) & " CON DNI No " & IIf(Len(rsDatosDDJJ!dniconyuge) > 0, Trim(rsDatosDDJJ!dniconyuge), IIf(Len(rsDatosDDJJ!dnicodeudor) > 0, Trim(rsDatosDDJJ!dnicodeudor), "**********")) & lnSaltoLinDoc
'            sCadImp = sCadImp & "DECLARO(AMOS) BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & lnSaltoLinDoc
'            sCadImp = sCadImp & "SIGUIENTES BIENES:" & lnSaltoLinDoc & lnSaltoLinDoc
'        Else
'            sCadImp = sCadImp & "LA EMPRESA, " & Trim(rsDatosDDJJ!cPersNombre) & ", CON RUC No " & Trim(rsDatosDDJJ!RUC) & lnSaltoLinDoc
'            sCadImp = sCadImp & "DECLARA BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & lnSaltoLinDoc
'            sCadImp = sCadImp & "SIGUIENTES BIENES:" & lnSaltoLinDoc & lnSaltoLinDoc
'        End If
'
'    sCadImp = sCadImp & "1.- BIENES INMUEBLES:" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                              AREA DE     AREA          TIPO                        VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  UBICACION                   TERRENO     CONSTRUIDA    D/USO         R.P.          DECLRADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsInmue.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cUbicacion, 23)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaTerreno, 10, 2, True)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaConstru, 10, 2, True)
'        sCadImp = sCadImp & Space(3) & oImpre.ImpreFormat(rsInmue!cTipoUso, 12)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cRRPP, 8)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsInmue!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsInmue.MoveNext
'        If rsInmue.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsInmue.Close
'    Set rsInmue = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
'
'    sCadImp = sCadImp & "2.- AUTOMOVILES Y VEHICULOS DE RECREACION:" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                                                                                    VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  MARCA                   MODELO            AÑO        PLACA                        ESTIMADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsVehic.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cMarca, 20)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cModelo, 10, 2, True)
'        sCadImp = sCadImp & Space(8) & oImpre.ImpreFormat(rsVehic!cFecFabrica, 10)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cPlaca, 10)
'        sCadImp = sCadImp & Space(12) & oImpre.ImpreFormat(rsVehic!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsVehic!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsVehic.MoveNext
'        If rsVehic.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsVehic.Close
'    Set rsVehic = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
'
'    sCadImp = sCadImp & "3.- OTROS (Artefactos, Maquinarias, Joyas, Acciones, etc.):" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                                                                                    VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  DESCRIPCION                                                                       DECLRADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsOtros.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsOtros!CDescripcion, 30)
'        sCadImp = sCadImp & Space(46) & oImpre.ImpreFormat(rsOtros!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsOtros!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsOtros.MoveNext
'        If rsOtros.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsOtros.Close
'    Set rsOtros = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    sCadImp = sCadImp & Space(40) & "TOTAL PATRIMONIO DECLARADO - " & oImpre.ImpreFormat(rsDatosDDJJ!MonedaPatri, 10) & " :" & Space(3) & oImpre.ImpreFormat(lnTotPatri, 10, 2, True) & lnSaltoLinDoc & lnSaltoLinDoc
'    sCadImp = sCadImp & "DEJO(AMOS) PLENA CONSTANCIA DE LOS BIENES DECLARADOS LO(LOS) OFREZCO (CEMOS) IRREVOCABLEMENTE  EN RESPALDO A  LOS" & lnSaltoLinDoc
'    sCadImp = sCadImp & "CREDITOS DIRECTOS E INDIRECTOS QUE TENGA A BIEN CONCEDERME(NOS) LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE  MAYNAS." & lnSaltoLinDoc
'    sCadImp = sCadImp & "DECLARO(MOS) QUE LA INFORMACION PROPORCIONADA EN ESTE DOCUMENTO ES AUTENTICA Y AUTORIZO(AMOS) A LA CAJA MUNICIPAL" & lnSaltoLinDoc
'    sCadImp = sCadImp & "A REALIZAR SU VERIFICACION. ESTE DOCUMENTO TIENE CARACTER DE DECLARACION JURADA DE ACUERDO CON EL ARTICULO 179 DE" & lnSaltoLinDoc
'    sCadImp = sCadImp & "LA LEY No 26702 DECLARO ASI MISMO QUE EN CASO DE PROPORCIONAR INFORMACION FALSA ESTARE COMPRENDIDO DENTRO  DE LOS" & lnSaltoLinDoc
'    sCadImp = sCadImp & "ALCANCES  DEL  ART. 247 DEL  CODIGO PENAL QUE  ESTABLECE LA TIPIFICACION  DEL DELITO  FINANCIERO POR PROPORCIONAR" & lnSaltoLinDoc
'    sCadImp = sCadImp & "INFORMACION FALSA A UNA EMPRESA DEL SISTEMA FINANCIERO." & lnSaltoLinDoc & lnSaltoLinDoc
'
'    sCadImp = sCadImp & Space(70) & Trim(rsAgencia!dist) & ", " & Format(pdFecSis, "dddd") & ", " & Day(pdFecSis) & " de " & Format(pdFecSis, "mmmm") & " de " & Year(pdFecSis) & lnSaltoLinDoc
'
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
'    If Len(rsDatosDDJJ!DNI) > 0 Then
'        If Len(rsDatosDDJJ!dniconyuge) > 0 Then
'            sCadImp = sCadImp & ".........................................." & Space(22) & ".........................................." & lnSaltoLinDoc
'            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & Space(2) & "CONYUGE: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!Conyuge), 50) & lnSaltoLinDoc
'            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & Space(44) & "DNI: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!dniconyuge), 15) & lnSaltoLinDoc
'        ElseIf Len(rsDatosDDJJ!dnicodeudor) > 0 Then
'            sCadImp = sCadImp & ".........................................." & Space(22) & ".........................................." & lnSaltoLinDoc
'            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & Space(2) & "CONYUGE: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!Codeudor), 50) & lnSaltoLinDoc
'            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & Space(44) & "DNI: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!dnicodeudor), 15) & lnSaltoLinDoc
'        Else
'            sCadImp = sCadImp & ".........................................." & lnSaltoLinDoc
'            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & lnSaltoLinDoc
'            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & lnSaltoLinDoc
'        End If
'    Else
'        sCadImp = sCadImp & ".........................................." & lnSaltoLinDoc
'        sCadImp = sCadImp & "DECLARANTE:" & Trim(rsDatosDDJJ!cPersNombre) & lnSaltoLinDoc
'        sCadImp = sCadImp & "RUC: " & Trim(rsDatosDDJJ!RUC) & lnSaltoLinDoc
'    End If
'
'    ImprimeDDJJPatrimonial = sCadImp
'    Screen.MousePointer = 0

'    Dim rsInmue As ADODB.Recordset
'    Dim rsOtros As ADODB.Recordset
'    Dim rsVehic As ADODB.Recordset
    Dim rsDatosDDJJ As ADODB.Recordset
    Dim rsAgencia As ADODB.Recordset
    Dim rsGarantTitular As ADODB.Recordset
    Dim rsGarantAvales As ADODB.Recordset
    Dim lnTotPatri As Double
    
    Dim oDAgencia As COMDConstantes.DCOMAgencias
    
    Dim oNCredito As COMNCredito.NCOMCredDoc
    Dim oDCredito As COMDCredito.DCOMCredDoc
    Dim oDPer As New COMDPersona.DCOMPersonas
    Dim oDCred As COMDCredito.DCOMCredito

    Dim sCadImp As String, lcTitulo As String, lcPersCod As String
    
    Dim lcDNI As String
    Dim lcDniconyuge As String, lcPersNombre As String, lcConyuge As String
    
    Dim i As Integer, J As Integer

    Screen.MousePointer = 11
       
    Set oDAgencia = New COMDConstantes.DCOMAgencias
        Set rsAgencia = oDAgencia.RecuperaAgencias(pcCodAge)
    Set oDAgencia = Nothing

    Set oDCredito = New COMDCredito.DCOMCredDoc
        Set rsDatosDDJJ = oDCredito.ObtenerDatosDDJJPatri(pcCuenta)
        Set rsGarantTitular = oDCredito.ObtenerGarantiasTitular(pcCuenta)
        Set rsGarantAvales = oDCredito.ObtenerGarantiasAvales(pcCuenta)
    Set oDCredito = Nothing
            
'    If rsDatosDDJJ.RecordCount = 0 Then
'        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
'        Screen.MousePointer = 0
'        Exit Function
'    End If


'lnSaltoPagina

    lnTotPatri = 0
    
'********************************-------------------- IMPRIME DDJJ TITULAY Y CODEUDOR

    sCadImp = ""
    'Call ImprimeCabeceraDocumento(sCadImp, "", pdFecSis, "", "DECLARACION JURADA PATRIMONIAL", "", , True, , "")
    sCadImp = sCadImp & "CAJA MAYNAS" & lnSaltoLinDoc
    sCadImp = sCadImp & "RUC : 20103845328" & lnSaltoLinDoc & lnSaltoLinDoc
    
    lcTitulo = "DECLARACION JURADA PATRIMONIAL"
    sCadImp = sCadImp & Space((EspHorxPag - Len(lcTitulo)) / 2 - 18) & lcTitulo & lnSaltoLinDoc
    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    If Len(rsDatosDDJJ!DNI) > 0 Then
        sCadImp = sCadImp & "YO, " & Trim(rsDatosDDJJ!cPersNombre) & ", IDENTIFICADO CON DNI No " & Trim(rsDatosDDJJ!DNI) & " DE ESTADO CIVIL, " & IIf(Len(rsDatosDDJJ!EstadoCiv) > 0, Trim(rsDatosDDJJ!EstadoCiv), "**********") & lnSaltoLinDoc
        sCadImp = sCadImp & "Y (nombre del conyugue, si fuese el caso) " & IIf(Len(rsDatosDDJJ!Conyuge) > 0, Trim(rsDatosDDJJ!Conyuge), IIf(Len(rsDatosDDJJ!Codeudor) > 0, Trim(rsDatosDDJJ!Codeudor), "**********")) & " CON DNI No " & IIf(Len(rsDatosDDJJ!dniconyuge) > 0, Trim(rsDatosDDJJ!dniconyuge), IIf(Len(rsDatosDDJJ!dnicodeudor) > 0, Trim(rsDatosDDJJ!dnicodeudor), "**********")) & lnSaltoLinDoc
        sCadImp = sCadImp & "DECLARO(AMOS) BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & lnSaltoLinDoc
        sCadImp = sCadImp & "SIGUIENTES BIENES:" & lnSaltoLinDoc & lnSaltoLinDoc
    Else
        sCadImp = sCadImp & "LA EMPRESA, " & Trim(rsDatosDDJJ!cPersNombre) & ", CON RUC No " & Trim(rsDatosDDJJ!RUC) & lnSaltoLinDoc
        sCadImp = sCadImp & "DECLARA BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & lnSaltoLinDoc
        sCadImp = sCadImp & "SIGUIENTES BIENES:" & lnSaltoLinDoc & lnSaltoLinDoc
    End If

'    sCadImp = sCadImp & "1.- BIENES INMUEBLES:" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                              AREA DE     AREA          TIPO                        VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  UBICACION                   TERRENO     CONSTRUIDA    D/USO         R.P.          DECLRADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsInmue.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cUbicacion, 23)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaTerreno, 10, 2, True)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaConstru, 10, 2, True)
'        sCadImp = sCadImp & Space(3) & oImpre.ImpreFormat(rsInmue!cTipoUso, 12)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cRRPP, 8)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsInmue!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsInmue.MoveNext
'        If rsInmue.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsInmue.Close
'    Set rsInmue = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
'
'    sCadImp = sCadImp & "2.- AUTOMOVILES Y VEHICULOS DE RECREACION:" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                                                                                    VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  MARCA                   MODELO            AÑO        PLACA                        ESTIMADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsVehic.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cMarca, 20)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cModelo, 10, 2, True)
'        sCadImp = sCadImp & Space(8) & oImpre.ImpreFormat(rsVehic!cFecFabrica, 10)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsVehic!cPlaca, 10)
'        sCadImp = sCadImp & Space(12) & oImpre.ImpreFormat(rsVehic!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsVehic!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsVehic.MoveNext
'        If rsVehic.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsVehic.Close
'    Set rsVehic = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
'
'    sCadImp = sCadImp & "3.- OTROS (Artefactos, Maquinarias, Joyas, Acciones, etc.):" & lnSaltoLinDoc
'    sCadImp = sCadImp & "                                                                                    VALOR   " & lnSaltoLinDoc
'    sCadImp = sCadImp & "  DESCRIPCION                                                                       DECLRADO" & lnSaltoLinDoc
'    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
'    Do While Not rsOtros.EOF
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsOtros!CDescripcion, 30)
'        sCadImp = sCadImp & Space(46) & oImpre.ImpreFormat(rsOtros!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsOtros!nValorCom
'        sCadImp = sCadImp & lnSaltoLinDoc
'        rsOtros.MoveNext
'        If rsOtros.EOF Then
'            Exit Do
'        End If
'    Loop
'    rsOtros.Close
'    Set rsOtros = Nothing
'    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc

    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & "                              AREA DE     AREA          TIPO                        VALOR   " & lnSaltoLinDoc
    sCadImp = sCadImp & "  UBICACION                   TERRENO     CONSTRUIDA    D/USO         R.P.          DECLRADO" & lnSaltoLinDoc
    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
    Do While Not rsGarantTitular.EOF
        sCadImp = sCadImp & oImpre.ImpreFormat(rsGarantTitular!CDescripcion, 30)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaTerreno, 10, 2, True)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaConstru, 10, 2, True)
'        sCadImp = sCadImp & Space(3) & oImpre.ImpreFormat(rsInmue!cTipoUso, 12)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cRRPP, 8)
'        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nValorCom, 10, 2, True)
'        lnTotPatri = lnTotPatri + rsInmue!nValorCom
        sCadImp = sCadImp & lnSaltoLinDoc
        rsGarantTitular.MoveNext
        If rsGarantTitular.EOF Then
            Exit Do
        End If
    Loop
    rsGarantTitular.Close
    Set rsGarantTitular = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc


    '----- pie de pagina
    sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(40) & "TOTAL PATRIMONIO DECLARADO - " & oImpre.ImpreFormat(rsDatosDDJJ!MonedaPatri, 10) & " :" & Space(3) & oImpre.ImpreFormat(lnTotPatri, 10, 2, True) & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & "DEJO(AMOS) PLENA CONSTANCIA DE LOS BIENES DECLARADOS LO(LOS) OFREZCO (CEMOS) IRREVOCABLEMENTE  EN RESPALDO A  LOS" & lnSaltoLinDoc
    sCadImp = sCadImp & "CREDITOS DIRECTOS E INDIRECTOS QUE TENGA A BIEN CONCEDERME(NOS) LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE  MAYNAS." & lnSaltoLinDoc
    sCadImp = sCadImp & "DECLARO(MOS) QUE LA INFORMACION PROPORCIONADA EN ESTE DOCUMENTO ES AUTENTICA Y AUTORIZO(AMOS) A LA CAJA MUNICIPAL" & lnSaltoLinDoc
    sCadImp = sCadImp & "A REALIZAR SU VERIFICACION. ESTE DOCUMENTO TIENE CARACTER DE DECLARACION JURADA DE ACUERDO CON EL ARTICULO 179 DE" & lnSaltoLinDoc
    sCadImp = sCadImp & "LA LEY No 26702 DECLARO ASI MISMO QUE EN CASO DE PROPORCIONAR INFORMACION FALSA ESTARE COMPRENDIDO DENTRO  DE LOS" & lnSaltoLinDoc
    sCadImp = sCadImp & "ALCANCES  DEL  ART. 247 DEL  CODIGO PENAL QUE  ESTABLECE LA TIPIFICACION  DEL DELITO  FINANCIERO POR PROPORCIONAR" & lnSaltoLinDoc
    sCadImp = sCadImp & "INFORMACION FALSA A UNA EMPRESA DEL SISTEMA FINANCIERO." & lnSaltoLinDoc & lnSaltoLinDoc
    
    sCadImp = sCadImp & Space(70) & Trim(rsAgencia!dist) & ", " & Format(pdFecSis, "dddd") & ", " & Day(pdFecSis) & " de " & Format(pdFecSis, "mmmm") & " de " & Year(pdFecSis) & lnSaltoLinDoc

    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    If Len(rsDatosDDJJ!DNI) > 0 Then
        If Len(rsDatosDDJJ!dniconyuge) > 0 Then
            sCadImp = sCadImp & ".........................................." & Space(22) & ".........................................." & lnSaltoLinDoc
            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & Space(2) & "CONYUGE: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!Conyuge), 50) & lnSaltoLinDoc
            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & Space(44) & "DNI: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!dniconyuge), 15) & lnSaltoLinDoc
        ElseIf Len(rsDatosDDJJ!dnicodeudor) > 0 Then
            sCadImp = sCadImp & ".........................................." & Space(22) & ".........................................." & lnSaltoLinDoc
            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & Space(2) & "CONYUGE: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!Codeudor), 50) & lnSaltoLinDoc
            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & Space(44) & "DNI: " & oImpre.ImpreFormat(Trim(rsDatosDDJJ!dnicodeudor), 15) & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & ".........................................." & lnSaltoLinDoc
            sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!cPersNombre), 50) & lnSaltoLinDoc
            sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(rsDatosDDJJ!DNI), 15) & lnSaltoLinDoc
        End If
    Else
        sCadImp = sCadImp & ".........................................." & lnSaltoLinDoc
        sCadImp = sCadImp & "DECLARANTE:" & Trim(rsDatosDDJJ!cPersNombre) & lnSaltoLinDoc
        sCadImp = sCadImp & "RUC: " & Trim(rsDatosDDJJ!RUC) & lnSaltoLinDoc
    End If
    '---------- fin pie de pagina
    
    
'*****************************--------------------- Garantias de los avales

If rsGarantAvales.RecordCount > 0 Then

    Do While Not rsGarantAvales.EOF
        sCadImp = sCadImp & lnSaltoPagina
        
        lcPersCod = rsGarantAvales!cPersCod
        lcDNI = rsGarantAvales!DNI
        lcDniconyuge = rsGarantAvales!dniconyuge
        lcPersNombre = rsGarantAvales!cPersNombre
        lcConyuge = rsGarantAvales!Conyuge
        
        sCadImp = sCadImp & "CAJA MAYNAS" & lnSaltoLinDoc
        sCadImp = sCadImp & "RUC : 20103845328" & lnSaltoLinDoc & lnSaltoLinDoc
        
        lcTitulo = "DECLARACION JURADA PATRIMONIAL"
        sCadImp = sCadImp & Space((EspHorxPag - Len(lcTitulo)) / 2 - 18) & lcTitulo & lnSaltoLinDoc
        sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    
        sCadImp = sCadImp & "YO, " & Trim(rsGarantAvales!cPersNombre) & ", IDENTIFICADO CON DNI No " & Trim(rsGarantAvales!DNI) & " DE ESTADO CIVIL, " & IIf(Len(rsGarantAvales!EstadoCiv) > 0, Trim(rsGarantAvales!EstadoCiv), "**********") & lnSaltoLinDoc
        sCadImp = sCadImp & "Y (nombre del conyugue, si fuese el caso) " & IIf(Len(rsGarantAvales!Conyuge) > 0, Trim(rsGarantAvales!Conyuge), "**********") & " CON DNI No " & IIf(Len(rsGarantAvales!dniconyuge) > 0, Trim(rsGarantAvales!dniconyuge), "**********") & lnSaltoLinDoc
        sCadImp = sCadImp & "DECLARO(AMOS) BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & lnSaltoLinDoc
        sCadImp = sCadImp & "SIGUIENTES BIENES:" & lnSaltoLinDoc & lnSaltoLinDoc
    
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & "                              AREA DE     AREA          TIPO                        VALOR   " & lnSaltoLinDoc
        sCadImp = sCadImp & "  UBICACION                   TERRENO     CONSTRUIDA    D/USO         R.P.          DECLRADO" & lnSaltoLinDoc
        sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
        
        Do While rsGarantAvales!cPersCod = lcPersCod
        
            sCadImp = sCadImp & oImpre.ImpreFormat(rsGarantAvales!CDescripcion, 30)
    '        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaTerreno, 10, 2, True)
    '        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nAreaConstru, 10, 2, True)
    '        sCadImp = sCadImp & Space(3) & oImpre.ImpreFormat(rsInmue!cTipoUso, 12)
    '        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!cRRPP, 8)
    '        sCadImp = sCadImp & oImpre.ImpreFormat(rsInmue!nValorCom, 10, 2, True)
    '        lnTotPatri = lnTotPatri + rsInmue!nValorCom
            sCadImp = sCadImp & lnSaltoLinDoc

            rsGarantAvales.MoveNext
            If rsGarantAvales.EOF Then
                Exit Do
            End If
        Loop
'        rsGarantTitular.Close
'        Set rsGarantTitular = Nothing
        sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    
        '----- pie de pagina
        sCadImp = sCadImp & String(100, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(40) & "TOTAL PATRIMONIO DECLARADO - :" & Space(3) & oImpre.ImpreFormat(lnTotPatri, 10, 2, True) & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & "DEJO(AMOS) PLENA CONSTANCIA DE LOS BIENES DECLARADOS LO(LOS) OFREZCO (CEMOS) IRREVOCABLEMENTE  EN RESPALDO A  LOS" & lnSaltoLinDoc
        sCadImp = sCadImp & "CREDITOS DIRECTOS E INDIRECTOS QUE TENGA A BIEN CONCEDERME(NOS) LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE  MAYNAS." & lnSaltoLinDoc
        sCadImp = sCadImp & "DECLARO(MOS) QUE LA INFORMACION PROPORCIONADA EN ESTE DOCUMENTO ES AUTENTICA Y AUTORIZO(AMOS) A LA CAJA MUNICIPAL" & lnSaltoLinDoc
        sCadImp = sCadImp & "A REALIZAR SU VERIFICACION. ESTE DOCUMENTO TIENE CARACTER DE DECLARACION JURADA DE ACUERDO CON EL ARTICULO 179 DE" & lnSaltoLinDoc
        sCadImp = sCadImp & "LA LEY No 26702 DECLARO ASI MISMO QUE EN CASO DE PROPORCIONAR INFORMACION FALSA ESTARE COMPRENDIDO DENTRO  DE LOS" & lnSaltoLinDoc
        sCadImp = sCadImp & "ALCANCES  DEL  ART. 247 DEL  CODIGO PENAL QUE  ESTABLECE LA TIPIFICACION  DEL DELITO  FINANCIERO POR PROPORCIONAR" & lnSaltoLinDoc
        sCadImp = sCadImp & "INFORMACION FALSA A UNA EMPRESA DEL SISTEMA FINANCIERO." & lnSaltoLinDoc & lnSaltoLinDoc

        sCadImp = sCadImp & Space(70) & Trim(rsAgencia!dist) & ", " & Format(pdFecSis, "dddd") & ", " & Day(pdFecSis) & " de " & Format(pdFecSis, "mmmm") & " de " & Year(pdFecSis) & lnSaltoLinDoc

        sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
            If Len(lcDniconyuge) > 0 Then
                sCadImp = sCadImp & ".........................................." & Space(22) & ".........................................." & lnSaltoLinDoc
                sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(lcPersNombre), 50) & Space(2) & "CONYUGE: " & oImpre.ImpreFormat(Trim(lcConyuge), 50) & lnSaltoLinDoc
                sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(lcDNI), 15) & Space(44) & "DNI: " & oImpre.ImpreFormat(Trim(lcDniconyuge), 15) & lnSaltoLinDoc
            Else
                sCadImp = sCadImp & ".........................................." & lnSaltoLinDoc
                sCadImp = sCadImp & "DECLARANTE:" & oImpre.ImpreFormat(Trim(lcPersNombre), 50) & lnSaltoLinDoc
                sCadImp = sCadImp & "DNI:" & oImpre.ImpreFormat(Trim(lcDNI), 15) & lnSaltoLinDoc
            End If
        '---------- fin pie de pagina

'        If rsGarantAvales.EOF Then
'            Exit Do
'        End If
    Loop
    rsGarantAvales.Close
    Set rsGarantAvales = Nothing
        

End If
    
    
    
    ImprimeDDJJPatrimonial = sCadImp
    Screen.MousePointer = 0

End Function

'*** PEAC 20080510
Public Function ImprimeListadoPolizas(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

'Dim oDCred As DCredDoc
Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim i As Integer
Dim sAgencia As String
Dim lcAge As String, lcCli As String

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosListadoPoliza(pdFecSis, pMatAgencias)
    Set oDCred = Nothing
    
    
'''-------------------------------------------------------------------------------------

    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONTROL DE POLIZAS AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadCab = sCadImp
    i = 7

    Do While Not R.EOF
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        lcAge = R!cCodAge
        Do While R!cCodAge = lcAge
            i = i + 1
            sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & oImpre.ImpreFormat("CUENTA", 18) & oImpre.ImpreFormat("M.DESEMB.", 10) & oImpre.ImpreFormat("SALDO CAP.", 9) & oImpre.ImpreFormat("MON.", 2) & oImpre.ImpreFormat("F.DESEMB.", 10) & oImpre.ImpreFormat("F.TERMINO", 10) & oImpre.ImpreFormat("POLIZA", 8) & oImpre.ImpreFormat("TIPO POLI.", 8) & oImpre.ImpreFormat("CIA.SEGUROS", 12) & oImpre.ImpreFormat("F.VIGENC.", 10) & oImpre.ImpreFormat("F.VENCIM", 10) & oImpre.ImpreFormat("CONDIC.", 8) & oImpre.ImpreFormat("DIRECC. INMUEBLE", 25) & oImpre.ImpreFormat("SUMA ASEG.$", 10) & lnSaltoLinDoc
            sCadImp = sCadImp & String(172, "-") & lnSaltoLinDoc
            lcCli = R!cPersCod
            Do While R!cPersCod = lcCli
                i = i + 1

                sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18) & oImpre.ImpreFormat(R!nMontoDesemb, 8, 2) & oImpre.ImpreFormat(R!nSaldoCap, 8, 2) & Space(1) & oImpre.ImpreFormat(R!cMoneda, 2) & oImpre.ImpreFormat(Format(R!dDesembolso, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(Format(R!dTermino, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(R!cCodPolizaAseg, 8) & oImpre.ImpreFormat(R!TipoPoliza, 8) & oImpre.ImpreFormat(R!cCiaSeguros, 12) & oImpre.ImpreFormat(Format(R!dVigenciaAseg, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(Format(R!dVencAseg, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(R!cCondicionPoliza, 8) & oImpre.ImpreFormat(R!cDireccionInmueble, 25) & oImpre.ImpreFormat(R!nSumaAsegurada, 8, 2) & lnSaltoLinDoc

                R.MoveNext

                If R.EOF Then
                    Exit Do
                End If

                If i >= 56 Then
                    sCadImp = sCadImp & lnSaltoPagina
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONTROL DE POLIZAS AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    'MAVM 26112009 ***
                    'sCadCab = sCadImp
                    '***
                    i = 7

                End If
            Loop
            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc

            If R.EOF Then
                Exit Do
            End If
        Loop

        If R.EOF Then
            Exit Do
        End If
    Loop


'''--------------------------------------------------

'    If R.RecordCount > 0 Then
'        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONTROL DE POLIZAS AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadCab = sCadImp
'        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'        sCadImp = sCadImp & oImpre.ImpreFormat("CUENTA", 18) & oImpre.ImpreFormat("M.DESEMB.", 10) & oImpre.ImpreFormat("SALDO CAP.", 9) & oImpre.ImpreFormat("MON.", 2) & oImpre.ImpreFormat("F.DESEMB.", 10) & oImpre.ImpreFormat("F.TERMINO", 10) & oImpre.ImpreFormat("POLIZA", 8) & oImpre.ImpreFormat("TIPO POLI.", 8) & oImpre.ImpreFormat("CIA.SEGUROS", 12) & oImpre.ImpreFormat("F.VIGENC.", 10) & oImpre.ImpreFormat("F.VENCIM", 10) & oImpre.ImpreFormat("CONDIC.", 8) & oImpre.ImpreFormat("DIRECC. INMUEBLE", 25) & oImpre.ImpreFormat("SUMA ASEG.$", 10) & lnSaltoLinDoc
'        sCadImp = sCadImp & String(172, "-") & lnSaltoLinDoc
'    End If
'
'    Do While Not R.EOF
'    'FINAL DE ANALISTA
'        If Trim(R!cCtaCod) = "TOTAL" Then
'            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'            sCadImp = sCadImp & lnSaltoLinDoc
'            R.MoveNext
'            If R.EOF Then
'                Exit Do
'            End If
'
'            'FINAL DE AGENCIA
'            If R!sCliente = "TOTAL" Then
'                sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'                sCadImp = sCadImp & lnSaltoLinDoc
'                R.MoveNext
'                If R.EOF Then
'                    Exit Do
'                End If
'                'TOTAL FINAL
'                If R!sAgencia = "TOTAL" Then
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    Exit Do
'                Else
'                    nPuntPag = 0
'                    sCadImp = sCadImp & Chr$(12)
'
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & oImpre.ImpreFormat("CUENTA", 20) & oImpre.ImpreFormat("M.DESEMB.", 10) & oImpre.ImpreFormat("SALDO CAP.", 10) & oImpre.ImpreFormat("MONEDA", 3) & oImpre.ImpreFormat("F.DESEMB.", 10) & oImpre.ImpreFormat("F.TERMINO", 10) & oImpre.ImpreFormat("POLIZA", 5) & oImpre.ImpreFormat("CIA.SEGUROS", 10) & oImpre.ImpreFormat("F.VIGENC.", 12) & oImpre.ImpreFormat("F.VENCIM", 12) & oImpre.ImpreFormat("CONDICION", 8) & oImpre.ImpreFormat("DIRECC. INMUEBLE", 25) & oImpre.ImpreFormat("SUMA ASEG.$", 12) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'                End If
'
'            Else
'                    sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & oImpre.ImpreFormat("CUENTA", 20) & oImpre.ImpreFormat("M.DESEMB.", 10) & oImpre.ImpreFormat("SALDO CAP.", 10) & oImpre.ImpreFormat("MONEDA", 3) & oImpre.ImpreFormat("F.DESEMB.", 10) & oImpre.ImpreFormat("F.TERMINO", 10) & oImpre.ImpreFormat("POLIZA", 5) & oImpre.ImpreFormat("CIA.SEGUROS", 10) & oImpre.ImpreFormat("F.VIGENC.", 12) & oImpre.ImpreFormat("F.VENCIM", 12) & oImpre.ImpreFormat("CONDICION", 8) & oImpre.ImpreFormat("DIRECC. INMUEBLE", 25) & oImpre.ImpreFormat("SUMA ASEG.$", 12) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
'            End If
'        End If
'
'        sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18) & oImpre.ImpreFormat(R!nMontoDesemb, 8, 2) & oImpre.ImpreFormat(R!nSaldoCap, 8, 2) & Space(1) & oImpre.ImpreFormat(R!cMoneda, 2) & oImpre.ImpreFormat(Format(R!dDesembolso, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(Format(R!dTermino, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(R!cCodPolizaAseg, 8) & oImpre.ImpreFormat(R!TipoPoliza, 8) & oImpre.ImpreFormat(R!cCiaSeguros, 12) & oImpre.ImpreFormat(Format(R!dVigenciaAseg, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(Format(R!dVencAseg, "dd/mm/yyyy"), 10) & oImpre.ImpreFormat(R!cCondicionPoliza, 8) & oImpre.ImpreFormat(R!cDireccionInmueble, 25) & oImpre.ImpreFormat(R!nSumaAsegurada, 8, 2) & lnSaltoLinDoc
'
'        R.MoveNext
'    Loop

''' ---------------------------------------------------------------------

    R.Close
    Set R = Nothing
    ImprimeListadoPolizas = sCadImp

End Function

'*** PEAC 20080627
Public Function nRepo108203_CreditosVigentes_Arqueo(ByVal nTipoReporte As Byte, ByVal cSubTit As String, ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal cMoneda As String, ByVal cProductos As String, ByVal cCondicion As String, ByVal psAnalistas, _
ByVal psAgencias As String, ByVal nProtesto As Integer, ByVal bOrdenCreditos As Boolean, _
Optional psCodAge As String, Optional psCodUser As String, Optional pdFecSis As Date, Optional psNomAge As String, Optional psNomCmac As String, Optional pnMontoIni As Double, Optional pnMontoFin As Double) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset


'Dim loDataRep As dColPFunciones
Dim oConec As COMConecta.DCOMConecta

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
'Dim loImprimeCab As NColPImpre
Dim sCliente As String * 35
Dim sMoneda As String * 7
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoAnalista As String
Dim i As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lnProtesto As Integer, lnOrdenCreditos As Integer '*** PEAC 20090406

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lsCondiciones = "Al " & pdFecFin

'**Modificado por DAOR 20070417, se agregó el join con la tabla ProductoTasaInteres
'**para obtener la tasa de interes compensatorio, además al cCtaCodAnt
'**en caso que sea vacio reemplazar por "9" para que se visualizen al ultimo

'---ARCV se modifico para mejorar el Performance no hacer (SELECT...) as Dato


'*** PEAC 20090406
'lsSQL = "SELECT PP.cCtaCod,case dbo.GetCtaAntigua(Producto.cCtaCod) when '' then '9' else dbo.GetCtaAntigua(Producto.cCtaCod) end as cCtaCodAnt,PP.cPersCod, dbo.Persona.cPersNombre, SUBSTRING(dbo.Producto.cCtaCod, 9, 1) AS cMoneda, " & _
            " Constante.cConsDescripcion AS cDesMoneda, Colocaciones.nMontoCol, Producto.nSaldo AS nSaldoCapital, Colocaciones.dVigencia, " & _
            " ANA.cPersCod AS cAnalista,PERAN.cPersNombre AS cNomAnalista,RH.cUser AS cUserAnalista,ColocacCred.nDiasAtraso,PT.nTasaInteres " & _
            " FROM ProductoPersona PP INNER JOIN Producto ON PP.cCtaCod = Producto.cCtaCod AND PP.nPrdPersRelac=20 INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante ON SUBSTRING(Producto.cCtaCod, 9, 1) = Constante.nConsValor INNER JOIN " & _
            " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod and Colocaciones.nMontoCol between " & pnMontoIni & " and " & pnMontoFin & " and Colocaciones.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' INNER JOIN " & _
            " ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
            " inner join ProductoTasaInteres PT on Producto.cCtaCod=PT.cCtaCod and PT.nPrdTasaInteres=" & gColocLineaCredTasasIntCompNormal & _
            "  " & _
            " INNER JOIN ProductoPersona ANA ON ANA.cCtaCod = Producto.cCtaCod AND ANA.nPrdPersRelac=28" & _
            " INNER JOIN Persona PERAN ON PERAN.cPersCod = ANA.cPersCod " & _
            " INNER JOIN RRHH RH ON RH.cPersCod = ANA.cPersCod " & _
            " WHERE (Producto.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
            " " '" AND (PP.nPrdPersRelac=" & gColRelPersTitular & ") "

'            If Len(Trim(cMoneda)) > 0 Then
'                lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & cMoneda & ")) "
'            End If
'
'            lsSQL = lsSQL & " AND (Constante.nConsCod=" & gMoneda & ") "

'            If Len(Trim(cCondicion)) > 0 Then
'                lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion IN (" & cCondicion & ")) "
'            End If

'            If Len(Trim(cProductos)) > 0 Then
'                lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & cProductos & "))"
'            End If

'            If Len(psAnalistas) > 0 Then
'            'ARCV 22-06-2007
'            '     lsSQL = lsSQL & " AND ((SELECT PRP.cPersCod FROM ProductoPersona PRP " & _
'                               " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & psAnalistas & ")) "
'                    lsSQL = lsSQL & " AND ANA.cPersCod IN (" & psAnalistas & ") "
'            '----------
'            End If

'            'asgnando las areas
'            lsSQL = lsSQL & " and Substring(Producto.cCtaCod,4,2) in ('" & psAgencias & "')"

            'asignado si existe o no protesto
'            If nProtesto = 1 Then
'                lsSQL = lsSQL & " and  Exists(Select *"
'                lsSQL = lsSQL & " From ColocCalendDet"
'                lsSQL = lsSQL & " Where cCtaCod=Producto.cCtaCod and nPrdConceptoCod in (1219,1220))"
'            ElseIf nProtesto = 2 Then
'                lsSQL = lsSQL & " and not Exists(Select *"
'                lsSQL = lsSQL & " From ColocCalendDet"
'                lsSQL = lsSQL & " Where cCtaCod=Producto.cCtaCod and nPrdConceptoCod in (1219,1220))"
'            End If

'        If nTipoReporte = 1 Then 'Ordenado por Codigo
'            If bOrdenCreditos = True Then
'                lsSQL = lsSQL & " ORDER BY PP.cCtaCod ASC "
'            Else
'                lsSQL = lsSQL & " ORDER BY cCtaCodAnt ASC "
'            End If
'            'lsSQL = lsSQL & " ORDER BY PP.cCtaCod ASC "
'        ElseIf nTipoReporte = 2 Then 'Agrupado por Analista
'            'Comentado por DAOR 20070418
'            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
'            '                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
'            '                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
'            '                " Persona.cPersNombre"
'
'            '**DAOR 20070418***********************************************
'            'ARCV ------- NO USAR (SELECT ....)
'            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
'                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
'                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") "
'            lsSQL = lsSQL & " ORDER BY PERAN.cPersNombre  "
'
'            If bOrdenCreditos = True Then
'                lsSQL = lsSQL & ", PP.cCtaCod ASC "
'            Else
'                lsSQL = lsSQL & ", cCtaCodAnt ASC "
'            End If
'            '**************************************************************
'        ElseIf nTipoReporte = 3 Then 'Analista Esp
'            'Comentado por DAOR 20070418
'            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
'            '                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
'            '                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
'            '                " Persona.cPersNombre"
'
'            '**DAOR 20070418***********************************************
'            'ARCV 22-06-2007
'            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
'                            " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
'                            " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") "
'            lsSQL = lsSQL & " ORDER BY PERAN.cPersNombre "
'            If bOrdenCreditos = True Then
'                lsSQL = lsSQL & ", PP.cCtaCod ASC "
'            Else
'                lsSQL = lsSQL & ", cCtaCodAnt ASC "
'            End If
'            '**************************************************************
'        End If


    '*** PEAC 20090406
    If Len(Trim(cMoneda)) = 0 Then
        cMoneda = "1,2"
    End If

    If bOrdenCreditos = True Then
        lnOrdenCreditos = 1 ''lsSQL = lsSQL & " ORDER BY PP.cCtaCod ASC "
    Else
        lnOrdenCreditos = 2  ''lsSQL = lsSQL & " ORDER BY cCtaCodAnt ASC "
    End If

    cMoneda = Replace(Replace(cMoneda, "'", ""), " ", "")
    cProductos = Replace(Replace(cProductos, "'", ""), " ", "")
    cCondicion = Replace(Replace(cCondicion, "'", ""), " ", "")
    cProductos = Replace(Replace(cProductos, "'", ""), " ", "")
    psAgencias = Replace(Replace(psAgencias, "'", ""), " ", "")
    psAnalistas = Replace(Replace(psAnalistas, "'", ""), " ", "")

    lsSQL = " exec stp_sel_ReporteCreditosVigentesArqueo " & _
                    pnMontoIni & "," & _
                    pnMontoFin & ",'" & _
                    Format(pdFecIni, "yyyymmdd") & "','" & _
                    Format(pdFecFin, "yyyymmdd") & "','" & _
                    cMoneda & "','" & _
                    cCondicion & "','" & _
                    cProductos & "','" & _
                    psAgencias & "'," & _
                    lnOrdenCreditos & ",'" & _
                    psAnalistas & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDataRep = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing



'**********************************************************

    'Dim sSql As String
    'Dim oConec As COMConecta.DCOMConecta
    'Dim rs As ADODB.Recordset
        
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    Set rs = oConec.CargaRecordSet(sSql)
'    oConec.CierraConexion
'    Set oConec = Nothing

'**********************************************************

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    Else
        nTNacionalAprob = 0
        nTNacionalSCapital = 0
        nTExtranjeroAprob = 0
        nTExtranjeroSCapital = 0
        lnLineas = 0
        lnPage = 1
        nItem = 0

        'CABECERA
        
        'lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)
        lsCadImp = lsCadImp
  
        'Format(pdFecIni, "yyyymmdd")
  
        Call ImprimeCabeceraDocumento(lsCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES - ARQUEO DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
                        'gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)
        lsCadImp = lsCadImp & " CREDITO            CREDITO ANTIGUO           CLIENTE                     MONEDA     M.APROBADO     S.CAPITAL DESEMBOLSO  TASA  ANA. ATRASO" & Chr(10)
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                        
        lnIndice = 5:  lnLineas = 5
         
        sTempoAnalista = IIf(IsNull(lrDataRep!cAnalista), "", lrDataRep!cAnalista)

        If nTipoReporte = 2 Or nTipoReporte = 3 Then
            'Analista
            lsCadImp = lsCadImp & lsNegritaOn & " ANALISTA         : " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
            lnLineas = lnLineas + 2
            lnIndice = lnIndice + 2
        End If

        With lrDataRep
            'RaiseEvent ShowProgress

            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
              
                'Totalizar por Analista
               If ((nTipoReporte = 2 Or nTipoReporte = 3) And sTempoAnalista <> lrDataRep!cAnalista) Then
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                    'Comentado por DAOR 20070418
                    'lsCadImp = lsCadImp & " TOTAL DE CASOS: " & oImpre.ImpreFormat(nitem, 5, 0) & Space(34) & "R E S U M E N   P O R   M O N E D A" & Chr(10) & Space(56) & "SOLES  " & Space(1) & oImpre.ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10) & _
                    '            Space(56) & "DOLARES" & Space(1) & oImpre.ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
                    
                    '**DAOR 20070418******************************************************
                    lsCadImp = lsCadImp & " TOTAL DE CASOS: " & oImpre.ImpreFormat(nItem, 5, 0) & Space(52) & "R E S U M E N   P O R   M O N E D A" & Chr(10)
                    lsCadImp = lsCadImp & Space(74) & "SOLES  " & Space(1) & oImpre.ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10)
                    lsCadImp = lsCadImp & Space(74) & "DOLARES" & Space(1) & oImpre.ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
                    '*********************************************************************

                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)

                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    
                    'lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)
                    
                    Call ImprimeCabeceraDocumento(lsCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES - ARQUEO DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
                    
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                    lsCadImp = lsCadImp & " CREDITO            CREDITO ANTIGUO           CLIENTE                     MONEDA     M.APROBADO     S.CAPITAL DESEMBOLSO  TASA  ANA. ATRASO" & Chr(10)
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)


                    lnLineas = 5
                    lnIndice = lnIndice + 5

                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & " ANALISTA         : " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    lnLineas = lnLineas + 2
                    lnIndice = lnIndice + 2
 
                    nItem = 0

                    nTNacionalAprob = 0
                    nTNacionalSCapital = 0
                    nTExtranjeroAprob = 0
                    nTExtranjeroSCapital = 0
                    sTempoAnalista = lrDataRep!cAnalista
                      
                    lnLineas = lnLineas + 6
                    lnIndice = lnIndice + 6
                    
                Else

                    If !cMoneda = gMonedaNacional Then
                        nTNacionalAprob = nTNacionalAprob + !nMontoCol
                        nTNacionalSCapital = nTNacionalSCapital + !nSaldoCapital
                    Else
                        nTExtranjeroAprob = nTExtranjeroAprob + !nMontoCol
                        nTExtranjeroSCapital = nTExtranjeroSCapital + !nSaldoCapital
                    End If

                End If
 
                nItem = nItem + 1

                sCliente = !cPersNombre
                sMoneda = !cDesMoneda
                
                'Comentado por DAOR 20070418
                'lsCadImp = lsCadImp & Space(1) & !cCtaCod & Space(1) & oImpre.ImpreFormat(IIf(IsNull(!CctacodAnt) Or !CctacodAnt = "", "NUEVO CREDITO", !CctacodAnt), 23) & sCliente & Space(1) & sMoneda & Space(1) & _
                '           oImpre.ImpreFormat(!nMontoCol, 10, 2) & Space(1) & oImpre.ImpreFormat(!nSaldoCapital, 10, 2) & Space(1) & _
                '           Format(!dVigencia, "dd/MM/YYYY") & Space(1) & !cUserAnalista & Space(1) & oImpre.ImpreFormat(!nDiasAtraso, 6, 0) & Chr(10)
                
                '**DAOR 20070418*********************************************************
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(!cCtaCod, 18, 0)
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(IIf(IsNull(!cCtaCodAnt) Or !cCtaCodAnt = "9", "NUEVO CREDITO", !cCtaCodAnt), 18, 0)
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(sCliente, 34, 0)
                lsCadImp = lsCadImp & Space(1) & sMoneda
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(!nMontoCol, 10, 2)
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(!nSaldoCapital, 10, 2)
                lsCadImp = lsCadImp & Space(1) & Format(!dVigencia, "dd/MM/YYYY")
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(!nTasaInteres, 2, 4)
                lsCadImp = lsCadImp & Space(1) & !cUserAnalista
                lsCadImp = lsCadImp & Space(1) & oImpre.ImpreFormat(!nDiasAtraso, 5, 0) & Chr(10)
                '************************************************************************


                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
 
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    'lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)
                    
                    Call ImprimeCabeceraDocumento(lsCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES - ARQUEO DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
                    
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                    lsCadImp = lsCadImp & " CREDITO            CREDITO ANTIGUO           CLIENTE                     MONEDA     M.APROBADO     S.CAPITAL DESEMBOLSO  TASA  ANA. ATRASO" & Chr(10)
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                    
                    
                    lnLineas = 5
                    lnIndice = lnIndice + 5
                End If
 
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
 
            Loop

            'RaiseEvent CloseProgress
 
        End With
 
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)

        'Comentado por DAOR
        'lsCadImp = lsCadImp & " TOTAL DE CASOS: " & oImpre.ImpreFormat(nitem, 5, 0) & Space(34) & "R E S U M E N   P O R   M O N E D A" & Chr(10) & Space(56) & "SOLES  " & Space(1) & oImpre.ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10) & _
        '        Space(56) & "DOLARES" & Space(1) & oImpre.ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
        
        '**DAOR 20070418******************************************************
        lsCadImp = lsCadImp & " TOTAL DE CASOS: " & oImpre.ImpreFormat(nItem, 5, 0) & Space(52) & "R E S U M E N   P O R   M O N E D A" & Chr(10)
        lsCadImp = lsCadImp & Space(74) & "SOLES  " & Space(1) & oImpre.ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10)
        lsCadImp = lsCadImp & Space(74) & "DOLARES" & Space(1) & oImpre.ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & oImpre.ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
        '*********************************************************************
                 
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)
 
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)

    End If

    nRepo108203_CreditosVigentes_Arqueo = lsCadBuffer

End Function


'*** PEAC 20080804 OPERACION 108350

Public Function ImprimeRepClientesHisNegativo(ByVal pMatAgencias As Variant, _
ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim sCadImp As String

Dim i As Integer, J As Integer
Dim sAgencia As String
Dim vage As String

    Screen.MousePointer = 11
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosClientesHisNegativo(pMatAgencias)
    
    Set oDCred = Nothing
    
    If R.BOF And R.EOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    sCadImp = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CLIENTES CON ANTECEDENTES NEGATIVOS (JUDICIAL/CASTIGADO)", psNomCmac, , True, , " AL " & Format(pdFecSis, "dd/mm/yyyy"))
        sCadImp = sCadImp & String(147, "-") & lnSaltoLinDoc
        '                    0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                              1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17
        sCadImp = sCadImp & "                                                                                  SALDO              TIPO          CONDIC.  ESTADO       FECHA PASE" & lnSaltoLinDoc
        sCadImp = sCadImp & "ITEM  CREDITO           No DOCUMENTO  COD.CLIENTE     NOMBRE CLIENTE              CAPITAL  ANALISTA  PRODUCTO      CREDITO  CREDITO      JUD/CAST  " & lnSaltoLinDoc
        sCadImp = sCadImp & String(147, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    i = 8
    J = 0
    Do While Not R.EOF
        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
        vage = R!CodAge
        
        J = 0
    
        Do While R!CodAge = vage
            i = i + 1
            
            J = J + 1
            sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!numdoc, 11)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cPersCod, 13)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cPersNombre, 22)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!nSaldo, 10, 2)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!cUser, 4)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!TipoProducto, 13)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!CondicionCredito, 8)
            sCadImp = sCadImp & oImpre.ImpreFormat(R!EstadoCredito, 12)
            sCadImp = sCadImp & oImpre.ImpreFormat(IIf(Year(R!FecPaseRecup) <= "1950", "  /  /    ", Format(R!FecPaseRecup, "dd/mm/yyyy")), 10)
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext

            If R.EOF Then
                Exit Do
            End If
        
            If i >= 56 Then
                sCadImp = sCadImp & lnSaltoPagina
                Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CLIENTES CON ANTECEDENTES NEGATIVOS (JUDICIAL/CASTIGADO)", psNomCmac, , True, , " AL " & Format(pdFecSis, "dd/mm/yyyy"))
                sCadImp = sCadImp & String(147, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & "                                                                                  SALDO              TIPO          CONDIC.  ESTADO       FECHA PASE" & lnSaltoLinDoc
                sCadImp = sCadImp & "ITEM  CREDITO           No DOCUMENTO  COD.CLIENTE     NOMBRE CLIENTE              CAPITAL  ANALISTA  PRODUCTO      CREDITO  CREDITO      JUD/CAST  " & lnSaltoLinDoc
                sCadImp = sCadImp & String(147, "-") & lnSaltoLinDoc & lnSaltoLinDoc
                i = 8
            End If
        
        Loop

    Loop

    R.Close
    Set R = Nothing
    ImprimeRepClientesHisNegativo = sCadImp
    Screen.MousePointer = 0
End Function
'ALPA 20110914********************
'Public Function ImprimeDocumentosDesembolsoL(ByVal psCtaCod As String, ByVal pnMovNro As Long, _
'    ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pdFecSis As Date, _
'    Optional ByVal pbApeAboCta As Integer, Optional ByVal psctaAho As String = "", _
'    Optional pnAhoMontoApe As Double = 0#, _
'    Optional pnIntGanadoAbono As Double = 0#, Optional pnIntGanadoRetGastos As Double = 0#, _
'    Optional pnIntGanadoRetCancel As Double = 0#, _
'    Optional pnMontoAbono As Double = 0#, Optional pnMontoRetGastos As Double = 0#, _
'    Optional pnMontoRetCancel As Double = 0#, _
'    Optional pnSaldoDispAhoAbono As Double = 0#, Optional ByVal pnSaldoCntAhoAbono As Double = 0#, _
'    Optional pnSaldoDispAhoRetGastos As Double = 0#, Optional pnSaldoCntAhoRetGastos As Double = 0#, _
'    Optional pnSaldoDispAhoRetCancel As Double = 0#, Optional pnSaldoCntAhoRetCancel As Double = 0#, _
'    Optional ByVal psNomCmac As String = "", Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, Optional pnTIF As Double = 0, _
'    Optional ByVal pnMontoPoliza As Double = 0) As Variant
'
'Dim MatDocs() As String
'Dim MatTempDesem(1, 2) As String
'Dim MatTempDesemGastos(1, 2) As String
'Dim MatTempDesemCancel() As String
'Dim MatTempComprob(1, 2) As String
'Dim MatTempPlan(2, 2) As String
'Dim MatTempAhoApe(1, 2) As String
'Dim MatTempAhoRetGastos(1, 2) As String
'Dim MatTempAhoRetCancel(1, 2) As String
'
'Dim nTotReg As Integer
'Dim oDCredDoc As COMDCredito.DCOMCredDoc
'Dim oDCredito As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim sCadTmp As String
'Dim AcumTmp As Double
'Dim i As Integer
'Dim nCapital As Double
'Dim nGasto As Double
'Dim nNumTran As Integer
'Dim nLin As Integer
'Dim nMontoDesemb As Double
'Dim nCalMiVivi As Boolean
'Dim dProxFecPag As Date
'
'    nTotReg = 0
'    '*****************************************************************************
'    '****   Recuperar Boletas de Desembolso
'    '*****************************************************************************
'
'    Set oDCredito = New COMDCredito.DCOMCredito
'    Set R = oDCredito.RecuperaDatosComunes(psCtaCod)
'    nCalMiVivi = R!bMiVivienda
'    nMontoDesemb = R!nMontoCol
'    R.Close
'    Set oDCredito = Nothing
'
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    Set R = oDCredDoc.RecuperaDatosBoletaDesembolso(pnMovNro, psCtaCod, psNomAge)
'    sCadTmp = ""
'    vsCadImpresion = ""
'    nNumTran = 0
'    Do While Not R.EOF
'        If R.Bookmark = 1 Then
'            nCapital = IIf(IsNull(R!nCapital), 0, R!nCapital)
'        Else
'            nGasto = IIf(IsNull(R!nGasto), 0, R!nGasto)
'        End If
'        nNumTran = nNumTran + R!nNumTran
'        R.MoveNext
'    Loop
'
'    R.MoveFirst
''    Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
'            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCmac)
'    'ARCV 30-07-2006
'    Call ImprimeCabeceraBoletaPago_NEW(psNomAge, R!sMoneda, "", _
'            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCMAC, R!cDNI)
'
'    'Cuerpo de la Boleta
'
'    sCadTmp = lsSaltoLin & lsSaltoLin & vsCadImpresion
'    sCadTmp = sCadTmp & lsNegritaOn
'
'If psCodCMAC = "102" Then  ' LIMA FEBM
'    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
'           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
'    Else
'            sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-") & lsSaltoLin
'    End If
'
'    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital, 12) & lsSaltoLin
'    'ARCV 30-01-2007
'    If Abs(nGasto) > 0 Then
'        sCadTmp = sCadTmp & "Gastos:" & oImpre.ImpreFormat(Abs(nGasto), 29) & lsSaltoLin
'    End If
'    '-------------
'    sCadTmp = sCadTmp & lsSaltoLin
'    sCadTmp = sCadTmp & lsSaltoLin
'    sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
'    sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto, 12) & lsSaltoLin
'    sCadTmp = sCadTmp & lsNegritaOff
'    sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 12) & lsSaltoLin
'    sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
'    sCadTmp = sCadTmp & lsSaltoLin
'    dProxFecPag = R!dVenc
'    R.Close
'
'Else   ' Liquidación Trujillo
'    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
'           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-")
'           sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
'    Else
'        sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-")
'        sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-") & lsSaltoLin
'    End If
'
'    'sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oimpre.ImpreFormat(nCapital + pnTIF, 12)
'    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital, 12)
'    'sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & oimpre.ImpreFormat(nCapital + pnTIF, 12) & lsSaltoLin
'    sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & oImpre.ImpreFormat(nCapital, 12) & lsSaltoLin
'    'ARCV 30-01-2007
'    If Abs(nGasto) > 0 Then
'        sCadTmp = sCadTmp & "Gastos:" & oImpre.ImpreFormat(Abs(nGasto), 29)
'        sCadTmp = sCadTmp & Space(14) & "Gastos :" & oImpre.ImpreFormat(Abs(nGasto), 28) & lsSaltoLin
'    End If
'    '-------
'
'    'ARCV 24-01-2007 + 30-01-2007
'    If Abs(pnMontoPoliza) > 0 Then
'    '    sCadTmp = sCadTmp & "Poliza:" & oImpre.ImpreFormat(Abs(pnMontoPoliza), 29)
'    '    sCadTmp = sCadTmp & Space(14) & "Poliza :" & oImpre.ImpreFormat(Abs(pnMontoPoliza), 28) & lsSaltoLin
'        sCadTmp = sCadTmp & "Poliza:" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnMontoPoliza), "#0.00") & "(-)", 15)
'        sCadTmp = sCadTmp & Space(2) & "Poliza :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnMontoPoliza), "#0.00") & "(-)", 28) & lsSaltoLin
'    End If
'    '---------------
'    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
'        sCadTmp = sCadTmp & lsSaltoLin
'    Else
'        sCadTmp = sCadTmp & "ITF   :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnTIF), "#0.00") & "(-)", 15)
'        sCadTmp = sCadTmp & Space(2) & "ITF    :" & String(26, " ") & oImpre.ImpreFormat(Format(Abs(pnTIF), "#.00") & "(-)", 28) & lsSaltoLin
'    End If
'
'    'ITF
'    'sCadTmp = sCadTmp & lsSaltoLin
'
'
'    sCadTmp = sCadTmp & lsSaltoLin
'    sCadTmp = sCadTmp & String(45, "-")
'    sCadTmp = sCadTmp & Space(7) & String(44, "-") & lsSaltoLin
'    'sCadTmp = sCadTmp & "Total a Desembolsar :" & Space(8) & oimpre.ImpreFormat(nCapital + NGasto, 12)
'    'sCadTmp = sCadTmp & Space(15) & "Total a Desembolsar  :" & Space(8) & oimpre.ImpreFormat(nCapital + NGasto, 12) & lsSaltoLin
'
'    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
'
'        'ARCV 24-01-2007
'        'sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto, 10)
'        'sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto, 12) & lsSaltoLin
'        sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 10)
'        sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnMontoPoliza, 12) & lsSaltoLin
'        '----------
'        sCadTmp = sCadTmp & lsNegritaOff
'
'    Else
'
'        'ARCV 24-01-2007
'        'sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnTIF, 10)
'        'sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnTIF, 12) & lsSaltoLin
'        sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnTIF - pnMontoPoliza, 10)
'        sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & oImpre.ImpreFormat(nCapital + nGasto - pnTIF - pnMontoPoliza, 12) & lsSaltoLin
'        '---------------------
'        sCadTmp = sCadTmp & lsNegritaOff
'    End If
'
'     sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 12)
'    'sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oimpre.ImpreFormat(nCapital + pnTIF, 12)
'    'sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oimpre.ImpreFormat(nCapital + pnTIF, 13) & lsSaltoLin
'    sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(nCapital, 13) & lsSaltoLin
'    sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
'    sCadTmp = sCadTmp & lsSaltoLin
'    dProxFecPag = R!dVenc
'    R.Close
'
'End If
' '' Fin
'
'    MatTempDesem(0, 0) = "1"
'    MatTempDesem(0, 1) = sCadTmp
'    nTotReg = nTotReg + 1
'
'    '*****************************************************************************
'    '********** Recuperar Boletas de Gastos de Desembolso
'    '*****************************************************************************
'    sCadTmp = ""
'    vsCadImpresion = ""
'    Set R = oDCredDoc.RecuperaDatosBoletaGastosDesembolso(pnMovNro, psCtaCod, psNomAge)
'    If Not R.BOF And Not R.EOF Then
'        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
'                R!FechaTran, R!HoraTran, R!cPersNombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC)
'        'Cuerpo de la Boleta
'
'        sCadTmp = lsSaltoLin & lsSaltoLin & vsCadImpresion
'        sCadTmp = sCadTmp & lsNegritaOn
'        If psCodCMAC = "102" Then
'            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-") & lsSaltoLin
'        Else
'            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-")
'            sCadTmp = sCadTmp & Space(7) & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(4, "-") & lsSaltoLin
'        End If
'    End If
'    AcumTmp = 0
'    Do While Not R.EOF
'     If psCodCMAC = 102 Then
'        sCadTmp = sCadTmp & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
'     Else
'        sCadTmp = sCadTmp & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12)
'        sCadTmp = sCadTmp & Space(11) & oImpre.ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & oImpre.ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
'     End If
'        AcumTmp = AcumTmp + Abs(R!nMonto)
'        R.MoveNext
'    Loop
'
'    If R.RecordCount > 0 Then
'        nLin = 7 - R.RecordCount
'        If nLin < 0 Then
'            nLin = 0
'        End If
'        For i = 1 To nLin
'            sCadTmp = sCadTmp & lsSaltoLin
'        Next i
'        If psCodCMAC = "102" Then
'            sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
'            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12) & lsSaltoLin
'            sCadTmp = sCadTmp & lsNegritaOff
'            R.MovePrevious
'            sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'            sCadTmp = sCadTmp & lsSaltoLin
'       Else
'            sCadTmp = sCadTmp & String(40, "-")
'            sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
'            sCadTmp = sCadTmp & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12)
'            sCadTmp = sCadTmp & Space(11) & oImpre.ImpreFormat("Total Gastos :", 25) & oImpre.ImpreFormat(AcumTmp, 12) & lsSaltoLin
'            sCadTmp = sCadTmp & lsNegritaOff
'            R.MovePrevious
'            sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'            sCadTmp = sCadTmp & lsSaltoLin
'        End If
'    End If
'    R.Close
'
'    MatTempDesemGastos(0, 0) = "1"
'    MatTempDesemGastos(0, 1) = sCadTmp
'    nTotReg = nTotReg + 1
'    '*****************************************************************************
'    '********** Recuperar Boletas de Cancelacion de Creditos
'    '*****************************************************************************
'    sCadTmp = ""
'    vsCadImpresion = ""
'    Set R = oDCredDoc.RecuperaDatosBoletaCancelCredDesembolso(pnMovNro, psCtaCod, psNomAge)
'    ReDim MatTempDesemCancel(R.RecordCount, 2)
'    Do While Not R.EOF
'    sCadTmp = ""
'    vsCadImpresion = ""
'    If psCodCMAC = "102" Then '' Liquidacion lima
'        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
'                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
'        'Cuerpo de la Boleta
'        sCadTmp = vsCadImpresion
'        sCadTmp = sCadTmp & lsNegritaOn
'        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
'
'        sCadTmp = sCadTmp & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
'        sCadTmp = sCadTmp & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
'        sCadTmp = sCadTmp & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
'        sCadTmp = sCadTmp & lsSaltoLin
'        sCadTmp = sCadTmp & lsSaltoLin
'        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
'        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
'        sCadTmp = sCadTmp & lsNegritaOff
'        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12) & lsSaltoLin
'        'sCadTmp = sCadTmp & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
'        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
'        nTotReg = nTotReg + 1
'        R.MoveNext
'    Else
'        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
'                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
'        'Cuerpo de la Boleta
'        sCadTmp = vsCadImpresion
'        sCadTmp = sCadTmp & lsNegritaOn
'        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-")
'        sCadTmp = sCadTmp & Space(14) & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
'
'        sCadTmp = sCadTmp & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12)
'        sCadTmp = sCadTmp & Space(16) & "Capital :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
'
'        sCadTmp = sCadTmp & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12)
'        sCadTmp = sCadTmp & Space(16) & "Interes :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
'
'        'ARCV 30-01-2007
'        If R!nGastos > 0 Then
'            sCadTmp = sCadTmp & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12)
'            sCadTmp = sCadTmp & Space(16) & "Gastos  :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
'        End If
'        '-------------
'        If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
'            sCadTmp = sCadTmp & lsSaltoLin
'        Else
'            sCadTmp = sCadTmp & "ITF     :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nITF), 0, R!nITF), 12)
'            sCadTmp = sCadTmp & Space(16) & "ITF     :" & Space(13) & oImpre.ImpreFormat(IIf(IsNull(R!nITF), 0, R!nITF), 12) & lsSaltoLin
'        End If
'
'
'        'sCadTmp = sCadTmp & lsSaltoLin
'        sCadTmp = sCadTmp & lsSaltoLin
'        sCadTmp = sCadTmp & String(40, "-")
'        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
'        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nITF), 0, R!nITF), 12)
'        sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & oImpre.ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nITF), 0, R!nITF), 12) & lsSaltoLin
'        sCadTmp = sCadTmp & lsNegritaOff
'        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12)
'        sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & oImpre.ImpreFormat(0, 12) & lsSaltoLin
'        'sCadTmp = sCadTmp & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & oimpre.ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        sCadTmp = sCadTmp & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & oImpre.ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
'        sCadTmp = sCadTmp & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin
'        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
'        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
'        nTotReg = nTotReg + 1
'        R.MoveNext
'    End If
'
'    Loop
'    R.Close
'
'    '***********************************************************
'    'Recuperar Boletas de Apertura de Cuenta
'    '***********************************************************
'    Dim oFecha As COMDConstSistema.DCOMGeneral
'    Set oFecha = New COMDConstSistema.DCOMGeneral
'    If pnAhoMontoApe > 0 Then
'        MatTempAhoApe(0, 0) = "1"
'        MatTempAhoApe(0, 1) = EmiteBoletaAhorro(psctaAho, pnAhoMontoApe, pnAhoMontoApe, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), psNomAge, psCodUsu, psCodCMAC, , pnTIF)
'        nTotReg = nTotReg + 1
'    End If
'
'
'    '***********************************************************
'    'Recuperar Boletas de Abono de Cuenta
'    '***********************************************************
'    If pnMontoAbono > 0 And pnAhoMontoApe <= 0 Then
'        MatTempAhoApe(0, 0) = "1"
'        MatTempAhoApe(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoAbono, pnSaldoDispAhoAbono, _
'            pnSaldoCntAhoAbono, pnIntGanadoAbono, 1, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
'            psNomAge, psCodUsu, psCodCMAC, pnTIF)
'        nTotReg = nTotReg + 1
'    End If
'
'    Dim C As COMDConstSistema.FCOMITF
'    Dim nItfValor As Double
'
'    Set C = New COMDConstSistema.FCOMITF
'    Call C.fgITFParametros
'    nItfValor = C.gnITFPorcent
'
'
'    '***********************************************************
'    'Recuperar Boletas de Retiro Gastos de Cuenta
'    '***********************************************************
'    If pnMontoRetGastos > 0 Then
'        MatTempAhoRetGastos(0, 0) = "1"
'        MatTempAhoRetGastos(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetGastos, pnSaldoDispAhoRetGastos, _
'            pnSaldoCntAhoRetGastos, pnIntGanadoRetGastos, 2, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
'            psNomAge, psCodUsu, psCodCMAC, C.fgITFCalculaImpuesto(pnMontoRetGastos))
'        nTotReg = nTotReg + 1
'    End If
'
'    '***********************************************************
'    'Recuperar Boletas de Retiro Cancelaciones de Cuenta
'    '***********************************************************
'    If pnMontoRetCancel > 0 Then
'        MatTempAhoRetCancel(0, 0) = "1"
'        MatTempAhoRetCancel(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetCancel, pnSaldoDispAhoRetCancel, _
'            pnSaldoCntAhoRetCancel, pnIntGanadoRetCancel, 2, Format(oFecha.FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
'            psNomAge, psCodUsu, psCodCMAC, C.fgITFCalculaImpuesto(pnMontoRetCancel))
'        nTotReg = nTotReg + 1
'    End If
'    Set oFecha = Nothing
'
'    Set C = Nothing
'
'    '********************************************************************************
'    'Recuperar Documento de Comprobante de Desembolso
'    '********************************************************************************
'    MatTempComprob(0, 0) = "2"
'    MatTempComprob(0, 1) = ImprimeComprobanteDesembolso(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, psNomCmac)
'    nTotReg = nTotReg + 1
'    '********************************************************************************
'    'Recuperar Documento de Plan de Pagos
'    '********************************************************************************
'    Dim loRep As COMNCredito.NCOMCalendario
'    Set loRep = New COMNCredito.NCOMCalendario
'    MatTempPlan(0, 0) = "2"
'    MatTempPlan(0, 1) = ""
'    'For I = 1 To 3
'    'ARCV 01-11-2006
'
'    '*** PEAC 20080723
'    'For i = 1 To 2
'        MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, , psNomCmac, , , True)
'        If nCalMiVivi = True Then
'            MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, True, , , , True)
'        End If
'    'Next i
'    '*** FIN PEAC
'
'    nTotReg = nTotReg + 1
'
'    ReDim MatDocs(nTotReg, 2)
'    MatDocs(0, 0) = MatTempDesem(0, 0)
'    MatDocs(0, 1) = MatTempDesem(0, 1)
'    MatDocs(1, 0) = MatTempDesemGastos(0, 0)
'    MatDocs(1, 1) = MatTempDesemGastos(0, 1)
'    nTotReg = 2
'    For i = 0 To UBound(MatTempDesemCancel) - 1
'        MatDocs(i + 2, 0) = MatTempDesemCancel(i, 0)
'        MatDocs(i + 2, 1) = MatTempDesemCancel(i, 1)
'        nTotReg = nTotReg + 1
'    Next i
'    'ARCV 01-04-2007   : Cambiar de Posicion
'    '    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
'    '    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
'    '    nTotReg = nTotReg + 1
'    '----------
'    'Apertura
'    If pnAhoMontoApe > 0 Then
'        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
'        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
'        nTotReg = nTotReg + 1
'    End If
'    'Abono
'    If pnMontoAbono > 0 And pnAhoMontoApe <= 0 Then
'        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
'        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
'        nTotReg = nTotReg + 1
'    End If
'    'Retiro de Gastos
'    If pnMontoRetGastos > 0 Then
'        MatDocs(nTotReg, 0) = MatTempAhoRetGastos(0, 0)
'        MatDocs(nTotReg, 1) = MatTempAhoRetGastos(0, 1)
'        nTotReg = nTotReg + 1
'    End If
'    'Retiro de Cancelaciones
'    If pnMontoRetCancel > 0 Then
'        MatDocs(nTotReg, 0) = MatTempAhoRetCancel(0, 0)
'        MatDocs(nTotReg, 1) = MatTempAhoRetCancel(0, 1)
'        nTotReg = nTotReg + 1
'    End If
'
'    'ARCV 01-04-2007: Cambiar de Posicion
'    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
'    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
'    nTotReg = nTotReg + 1
'    '------------
'
'    MatDocs(nTotReg, 0) = MatTempPlan(0, 0)
'    MatDocs(nTotReg, 1) = MatTempPlan(0, 1)
'
'    ImprimeDocumentosDesembolsoL = MatDocs
'
'End Function
'
'JUEZ 20130413 ************************************************************************************************
Public Function ImprimeBoletaComision(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psDOI As String, ByVal psCodCta As String, Optional ByVal pbCuenta As Boolean = False, _
        Optional ByVal psNumDoc As String, Optional ByVal psNomDoc As String, Optional ByVal psTexto As String = "XXX", Optional ByVal pdFecSis As Date, _
        Optional ByVal sNomAge As String = "", Optional ByVal sCodUser As String = "", Optional ByVal sLpt As String = "", _
        Optional ByVal bCadena As Boolean = False, Optional ByVal pbImpTMU As Boolean = False, Optional ByVal pbComAho As Boolean = False, _
        Optional ByVal psTipoPago As String = "", Optional ByVal pnITF As String = "", Optional psCtaCodCargoCta As String = "") As String 'JUEZ 20130902 Se agregó pbComAho,pnTipoPago,pnITF

    Dim nFicSal As Integer
    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sSDisp As String
    Dim sMax As Integer
    Dim saux As Integer
    Dim lsNroExt As String
    Dim lnTope As Integer
    Dim lsCadArg As String
    Dim lsInteres As String
    Dim sCad As String
    Dim lnCliAux As Long
    Dim lsCliAux1 As String
    Dim lsCliAux2 As String
    Dim sMontoTotal As String

    Dim lnChq As Long
    Dim lsChqAux1 As String
    Dim lsChqAux2 As String
    Dim lsNomAge As String
    Dim psMonedaTMU As String
    Dim oImpre As New COMFunciones.FCOMImpresion

ETIQ:

    On Error GoTo Error

    sCad = sCad & Chr$(27) & Chr$(64)
    sCad = sCad & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
    sCad = sCad & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
    sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    sCad = sCad & Chr$(27) + Chr$(18)  ' cancela condensada
    sCad = sCad & Chr$(27) + Chr$(72)  ' desactiva negrita

    sSep = 15
    sIni = 1
    sMax = 33
    saux = 5

    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")
    sMontoTotal = Format(CDbl(sMonto) + CDbl(IIf(pnITF = "", 0, pnITF)), "#,##0.00")

    lsNomAge = sNomAge

    'sCad = sCad & lsNegritaOn 'Activa Negrita

    If pbImpTMU = False Then
        sCad = sCad & lsSaltoLin & lsNegritaOn 'Activa Negrita
        sCad = sCad & Space(sIni) & "CMAC-MAYNAS" & IIf(pbComAho, Space(13), " - COMISIONES") & Space(9 + sSep + saux) & "CMAC-MAYNAS" & IIf(pbComAho, Space(13), " - COMISIONES") & lsSaltoLin 'JUEZ 20130902
    
        If Mid(psCodCta, 3, 3) = "000" Then
            If Mid(psCodCta, 6, 1) = 1 Then
                psMonedaTMU = "1"
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & lsSaltoLin
            Else
                psMonedaTMU = "2"
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & lsSaltoLin
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                psMonedaTMU = "1"
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & lsSaltoLin
            Else
                psMonedaTMU = "2"
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & lsSaltoLin
            End If
        End If
            
        sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & lsSaltoLin
      
        If psCliente <> "" Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & lsSaltoLin
        End If
        If psDOI <> "" Then
            sCad = sCad & " DOI:" & Space(4) & psDOI & Space(45 - Len(psDOI)) & "DOI:" & Space(4) & psDOI & lsSaltoLin
        End If
        If pbComAho = True And Len(psCodCta) = 18 Then
            sCad = sCad & " Cuenta: " & psCodCta & Space(45 - Len(psCodCta)) & "Cuenta: " & psCodCta & lsSaltoLin
        End If
        If pbComAho Then
            sCad = sCad & " TIPO:" & Space(3) & psTipoPago & Space(45 - Len(psTipoPago)) & "TIPO:" & Space(3) & psTipoPago & lsSaltoLin
        End If
        If psCliente <> "" And psDOI <> "" Then
            sCad = sCad & lsSaltoLin
        End If
        If psNumDoc <> "" Then
            sCad = sCad & lsSaltoLin
            sCad = sCad & Space(sIni) & psNomDoc & Space(sMax + 6 - Len(psNomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & psNomDoc & Space(sMax + 6 - Len(psNomDoc) - Len(psNumDoc)) & psNumDoc & lsSaltoLin
        End If

        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        'psTit = IF psTit
        sCad = sCad & lsNegritaOn 'Activa Negrita
        sCad = sCad & Space(sIni) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & lsSaltoLin
    
        lnChq = InStr(1, pstext, "*", vbTextCompare)
    
        If pnITF <> "" And pnITF <> "0.00" Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp("ITF Efectivo") & Space(sMax + 6 - Len("ITF Efectivo") - Len(pnITF)) & pnITF & Space(-1 + sSep) & oImpre.ImpreCarEsp("ITF Efectivo") & Space(sMax + 6 - Len("ITF Efectivo") - Len(pnITF)) & pnITF & lsSaltoLin
        Else
            sCad = sCad & lsSaltoLin
        End If
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & lsSaltoLin
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & lsSaltoLin
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & lsSaltoLin
            End If
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & lsSaltoLin
        End If
    
        sCad = sCad & lsNegritaOff 'Desactiva Negrita
        
        sCad = sCad & Space(sIni) & String(39, "-") & Space(-1 + sSep) & String(39, "-") & lsSaltoLin
        sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp("Total a pagar:") & Space(sMax + 6 - Len("Total a pagar:") - Len(sMontoTotal)) & sMontoTotal & Space(-1 + sSep) & oImpre.ImpreCarEsp("Total a pagar:") & Space(sMax + 6 - Len("Total a pagar:") - Len(sMontoTotal)) & sMontoTotal & lsSaltoLin & lsSaltoLin
        
        If Len(Trim(psCtaCodCargoCta)) = 18 Then
                sCad = sCad & Left(" Datos de cuenta origen:" & Space(52), 40) & Space(14) & Left("Datos de cuenta origen:" & Space(52), 40) & lsSaltoLin
                sCad = sCad & Left(" Nro. de cuenta    :" & Space(25), 20) & Right(Space(35) & Trim(psCtaCodCargoCta), 20) & Space(14) & Left("Nro. de cuenta     :" & Space(25), 20) & Right(Space(35) & Trim(psCtaCodCargoCta), 20) & lsSaltoLin
                sCad = sCad & Left(" Moneda            :" & Space(25), 20) & Right(Space(35) & Trim(IIf(psMonedaTMU = "1", "SOLES", "DOLARES")), 20) & Space(14) & Left("Moneda             :" & Space(25), 20) & Right(Space(35) & Trim(IIf(psMonedaTMU = "1", "SOLES", "DOLARES")), 20) & lsSaltoLin
                sCad = sCad & Left(" Monto cargado     : " & Space(25), 20) & Space(15) & Format$(sMontoTotal, "#,##0.00") & Space(14) & _
                Left("Monto cargado      : " & Space(25), 20) & Space(15) & Format$(sMontoTotal, "#,##0.00") & Space(40) & lsSaltoLin & lsSaltoLin
        End If
        
        sCad = sCad & " ** ESTIMADO CLIENTE, ANTES DE RETIRARSE" & Space(14) & "** ESTIMADO CLIENTE, ANTES DE RETIRARSE" & lsSaltoLin
        sCad = sCad & " DE  LA  VENTANILLA  VERIFIQUE  QUE  LOS" & Space(14) & "DE  LA  VENTANILLA  VERIFIQUE  QUE  LOS" & lsSaltoLin
        sCad = sCad & " DATOS  CONSIGNADOS  CORRESPONDAN  A  LA" & Space(14) & "DATOS  CONSIGNADOS  CORRESPONDAN  A  LA" & lsSaltoLin
        sCad = sCad & " OPERACION REALIZADA **                 " & Space(14) & "OPERACION REALIZADA **                 " & lsSaltoLin
        
        sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & lsSaltoLin
    Else
        sCad = sCad & lsSaltoLin & lsNegritaOn 'Activa Negrita
        sCad = sCad & "CMAC-MAYNAS" & IIf(pbComAho, "", " - COMISIONES") & lsSaltoLin 'JUEZ 20130902
        If Mid(psCodCta, 3, 3) = "000" Then
            If Mid(psCodCta, 6, 1) = 1 Then
                psMonedaTMU = "1"
                sCad = sCad & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & lsSaltoLin
            Else
                psMonedaTMU = "2"
                sCad = sCad & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & Space(sSep) & lsSaltoLin
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                psMonedaTMU = "1"
                sCad = sCad & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-SOLES")) & lsSaltoLin
            Else
                psMonedaTMU = "2"
                sCad = sCad & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len("-DOLARES")) & Space(sSep) & lsSaltoLin
            End If
        End If
        sCad = sCad & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & lsSaltoLin
        
        If psCliente <> "" Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            sCad = sCad & oImpre.ImpreCarEsp(psCliente) & lsSaltoLin
        End If
        If psDOI <> "" Then
            sCad = sCad & "DOI:" & Space(4) & psDOI & lsSaltoLin
        End If
        If pbComAho = True And Len(psCodCta) = 18 Then
            sCad = sCad & " Cuenta: " & psCodCta & lsSaltoLin
        End If
        If pbComAho Then
            sCad = sCad & "TIPO:" & Space(3) & psTipoPago & lsSaltoLin
        End If
        If psCliente <> "" And psDOI <> "" Then
            sCad = sCad & lsSaltoLin
        End If
        If psNomDoc <> "" Then
            sCad = sCad & lsSaltoLin
            sCad = sCad & psNomDoc & Space(sMax + 6 - Len(psNomDoc) - Len(psNumDoc)) & psNumDoc & lsSaltoLin
        End If

        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        sCad = sCad & lsNegritaOn 'Activa Negrita
        sCad = sCad & "-----" & psTit & "-----" & lsSaltoLin
    
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If pnITF <> "" And pnITF <> "0.00" Then
            pnITF = Format(pnITF, "#,##0.00")
            sCad = sCad & oImpre.ImpreCarEsp("ITF Efectivo" & Space(sMax + 6 - Len("ITF Efectivo") - Len(pnITF))) & pnITF & lsSaltoLin
        Else
            sCad = sCad & lsSaltoLin
        End If
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                sCad = sCad & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & lsSaltoLin
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                sCad = sCad & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & lsSaltoLin
                sCad = sCad & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & lsSaltoLin
            End If
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & lsSaltoLin
        End If
    
        sCad = sCad & lsNegritaOff 'Desactiva Negrita
    
        sCad = sCad & String(39, "-") & lsSaltoLin
        sCad = sCad & oImpre.ImpreCarEsp("Total a pagar:") & Space(sMax + 6 - Len("Total a pagar:") - Len(sMontoTotal)) & sMontoTotal & lsSaltoLin & lsSaltoLin
        
        If Len(Trim(psCtaCodCargoCta)) = 18 Then
            sCad = sCad & "Datos de cuenta origen:" & lsSaltoLin
            sCad = sCad & Left("Nro. de cuenta     :" & Space(25), 21) & Right(Space(35) & Trim(psCtaCodCargoCta), 18) & lsSaltoLin
            sCad = sCad & Left("Moneda             :" & Space(25), 21) & Right(Space(35) & Trim(IIf(psMonedaTMU = "1", "SOLES", "DOLARES")), 18) & lsSaltoLin
            sCad = sCad & Left("Monto cargado      : " & Space(18), 21) & Space(13) & Format$(sMontoTotal, "#,##0.00") & Space(35) & lsSaltoLin & lsSaltoLin 'MODIFICO APRI 20170426
        End If
        
        sCad = sCad & "** ESTIMADO CLIENTE, ANTES DE RETIRARSE" & lsSaltoLin
        sCad = sCad & "DE  LA  VENTANILLA  VERIFIQUE  QUE  LOS" & lsSaltoLin
        sCad = sCad & "DATOS  CONSIGNADOS  CORRESPONDAN  A  LA" & lsSaltoLin
        sCad = sCad & "OPERACION REALIZADA **                 " & lsSaltoLin
        
        sCad = sCad & oImpre.ImpreCarEsp(sCodUser) & lsSaltoLin
        sCad = sCad & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin
        sCad = sCad & lsSaltoLin & lsSaltoLin & lsSaltoLin & Chr(27) & "m" & Chr(10)
    End If

    ImprimeBoletaComision = sCad

    Exit Function
Error:
   
    Err.Raise Err.Number, "Error al realizar la Impresión...", Err.Description
End Function
'END JUEZ *****************************************************************************************************

' RIRO20131202 01-ERS098
Function HistorialPagoCreditos(ByVal psCtaCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    On Error GoTo Error
    
    sSql = "stp_sel_HistorialCredito '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Set HistorialPagoCreditos = rs
    Set rs = Nothing
    
 Exit Function

Error:
   
    Err.Raise Err.Number, "Error al obtener el historial del credito...", Err.Description
    
End Function
' END RIRO

'WIOR 20131226 ***************************************************************************************
Public Function ImprimePagoHonramiento(ByVal psCtaCod As String, ByVal psNombreCliente As String, ByVal psNomAge As String, _
                                        ByVal pdFecha As Date, ByVal pnMonto As Double, ByVal psNomCmacAbrev As String, _
                                        ByVal pnITF As Double, ByVal psCodCMAC As String, ByVal psCodUsu As String, Optional ByVal pbImpTMU As Boolean = False) As String
Dim oFecha As New COMDConstSistema.DCOMGeneral

Dim lsMoneda As String
lsMoneda = IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES")

vsCadImpresion = ""
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
vsCadImpresion = vsCadImpresion & lsSaltoLin

If pbImpTMU Then
    vsCadImpresion = Replace(vsCadImpresion, "k", "")
End If

Call ImprimeCabeceraBoletaPago(psNomAge, lsMoneda, 0, pdFecha, Format(oFecha.FechaHora(pdFecha), "hh:mm:ss"), psNombreCliente, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC, pbImpTMU, True)

'Cuerpo de la Boleta
vsCadImpresion = vsCadImpresion & lsNegritaOn

If pbImpTMU Then
    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO DE HONRAMIENTO " & String(10, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & "Monto Pagado :" & Space(11) & oImpre.ImpreFormat(pnMonto, 12) & lsSaltoLin
    
    If pnITF > 0 Then
        vsCadImpresion = vsCadImpresion & "ITF          :" & Space(11) & oImpre.ImpreFormat(pnITF, 12) & lsSaltoLin
    End If
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(11) & oImpre.ImpreFormat(pnMonto + pnITF, 12) & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psCodUsu, 51, 0) & lsSaltoLin & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin & lsSaltoLin & Chr(27) & "m" & Chr(10)
Else
    vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO DE HONRAMIENTO " & String(10, "-")
    vsCadImpresion = vsCadImpresion & Space(10) & String(10, "-") & " PAGO DE HONRAMIENTO " & String(10, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & "Monto Pagado :" & Space(12) & oImpre.ImpreFormat(pnMonto, 12)
    vsCadImpresion = vsCadImpresion & Space(10) & "Monto Pagado :" & Space(12) & oImpre.ImpreFormat(pnMonto, 12) & lsSaltoLin
    
    If pnITF > 0 Then
        vsCadImpresion = vsCadImpresion & "ITF          :" & Space(12) & oImpre.ImpreFormat(pnITF, 12)
        vsCadImpresion = vsCadImpresion & Space(10) & "ITF          :" & Space(12) & oImpre.ImpreFormat(pnITF, 12) & lsSaltoLin
    End If
    
    vsCadImpresion = vsCadImpresion & String(41, "-") & Space(10) & String(41, "-") & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(12) & oImpre.ImpreFormat(pnMonto + pnITF, 12)
    vsCadImpresion = vsCadImpresion & Space(10) & "Total Pagado :" & Space(12) & oImpre.ImpreFormat(pnMonto + pnITF, 12) & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(psCodUsu, 51, 0) & oImpre.ImpreFormat(psCodUsu, 53, 0) & lsSaltoLin & lsSaltoLin
    
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
End If
Set oFecha = Nothing
ImprimePagoHonramiento = vsCadImpresion
End Function
'WIOR FIN ********************************************************************************************
'ALPA 20160616****************************************************************************************
 Public Function ImprimeCuotasEndeudamiento(ByVal psCtaCod As String, ByVal psNombreCliente As String, ByVal psNomAge As String, _
                                        ByVal pdFecha As Date, ByVal pnMonto As Double, ByVal psNomCmacAbrev As String, _
                                        ByVal pnITF As Double, ByVal psCodCMAC As String, ByVal psCodUsu As String, Optional ByVal pbImpTMU As Boolean = False) As String
Dim oFecha As New COMDConstSistema.DCOMGeneral

Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim ldFecha As Date
Dim i As Integer
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Set oRs = objCredito.RecuperaCredEndeuCuotaSistFinanc(psCtaCod)
If Not (oRs.BOF Or oRs.EOF) Then
    ldFecha = oRs!dFecha
    oRs.MoveFirst
End If


Dim lsMoneda As String
lsMoneda = IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES")

vsCadImpresion = ""
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
vsCadImpresion = vsCadImpresion & lsSaltoLin

Call ImprimeCabeceraDocumento(vsCadImpresion, psNomAge, pdFecha, psCodUsu, "ENDEUDAMIENTO EN EL SISTEMA FINANCIERO/CUOTAS", psNomCmacAbrev, , True, , " AL " & Format(ldFecha, "dd/mm/yyyy"))
    vsCadImpresion = vsCadImpresion & String(160, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CLIENTE:::" & psNombreCliente & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CREDITO:::" & psCtaCod & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(160, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CODIGO", 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("EMPR.FINAN.", 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DIA DE PAGO", 10) & Space(10)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DEUDA", 12) & Space(10)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CUOTA", 12) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DEVENGADO", 12) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("OBSERVACION", 30) & Space(5)
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(160, "-") & lsSaltoLin

Do While Not oRs.EOF
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Cod_Emp, 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Nombre, 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(Right("0" & CStr(oRs!nDiasPago), 2), 10) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!nDeudaIFI, 12) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!nCuotaEstimada, 12) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!nInteresDevengado, 12) & Space(5)
    If oRs!nDeudaIFI <= 0 Then
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("Solicitar cronograma de pago", 30) & Space(5)
    Else
    vsCadImpresion = vsCadImpresion & String(30, " ") & Space(5)
    End If
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    oRs.MoveNext
Loop
vsCadImpresion = vsCadImpresion & String(160, "-") & lsSaltoLin
Set oFecha = Nothing
ImprimeCuotasEndeudamiento = vsCadImpresion
End Function
 Public Function ImprimeCredEndeuLineaNoUtilizada(ByVal psPersCod As String, ByVal psCtaCod As String, ByVal psNombreCliente As String, ByVal psNomAge As String, _
                                        ByVal pdFecha As Date, ByVal pnMonto As Double, ByVal psNomCmacAbrev As String, _
                                        ByVal pnITF As Double, ByVal psCodCMAC As String, ByVal psCodUsu As String, Optional ByVal pbImpTMU As Boolean = False) As String
Dim oFecha As New COMDConstSistema.DCOMGeneral

Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim lnValorDeuda As Currency
Dim ldFecha As Date
Dim i As Integer
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Set oRs = objCredito.RecuperaCredEndeuLineaNoUtilizada(psPersCod, psCtaCod)
If Not (oRs.BOF Or oRs.EOF) Then
    ldFecha = oRs!dFecha
    oRs.MoveFirst
End If


Dim lsMoneda As String
lsMoneda = IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES")

vsCadImpresion = ""
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
vsCadImpresion = vsCadImpresion & lsSaltoLin

Call ImprimeCabeceraDocumento(vsCadImpresion, psNomAge, pdFecha, psCodUsu, "ENDEUDAMIENTO EN LINEAS (TARJETAS) NO UTILIZADA", psNomCmacAbrev, , True, , " AL " & Format(ldFecha, "dd/mm/yyyy"))
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CLIENTE:::" & psNombreCliente & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CREDITO:::" & psCtaCod & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CODIGO", 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("EMPR.FINAN.", 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DEUDA NO UTILIZADA", 20) & Space(10)
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    lnValorDeuda = 0

Do While Not oRs.EOF
    If oRs!bActivo = 1 Then
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Cod_Emp, 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Nombre, 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Val_Saldo, 12) & Space(5)
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    lnValorDeuda = lnValorDeuda + IIf(IsNull(oRs!Val_Saldo), 0, oRs!Val_Saldo)
    End If
    oRs.MoveNext
Loop
vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
vsCadImpresion = vsCadImpresion & Space(22) & oImpre.ImpreFormat("CUOTA ESPERADA", 20) & Space(5)
vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(objCredito.ObtenerCuotaEsperadaLC(lnValorDeuda), 12) & Space(5)
Set oFecha = Nothing
ImprimeCredEndeuLineaNoUtilizada = vsCadImpresion
End Function
'FIN ALPA*********************************************************************************************
'CTI3 ERS032020***************************************************************************************
Public Function ImprimeCredDeudaSistemaFinanciero(ByVal psPersCod As String, ByVal psCtaCod As String, ByVal psNombreCliente As String, ByVal psNomAge As String, _
                                        ByVal pdFecha As Date, ByVal pnMonto As Double, ByVal psNomCmacAbrev As String, _
                                        ByVal pnITF As Double, ByVal psCodCMAC As String, ByVal psCodUsu As String, Optional ByVal pbImpTMU As Boolean = False) As String
Dim oFecha As New COMDConstSistema.DCOMGeneral

Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim lnValorDeuda As Currency
Dim ldFecha As Date
Dim i As Integer
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Set oRs = objCredito.RecuperaCredDeudaSistemaFinanciero(psPersCod, psCtaCod)
If Not (oRs.BOF Or oRs.EOF) Then
    ldFecha = oRs!dFecha
    oRs.MoveFirst
End If


Dim lsMoneda As String
lsMoneda = IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES")

vsCadImpresion = ""
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
vsCadImpresion = vsCadImpresion & lsSaltoLin

Call ImprimeCabeceraDocumento(vsCadImpresion, psNomAge, pdFecha, psCodUsu, "ENDEUDAMIENTO EN EL SISTEMA FINANCIERO", psNomCmacAbrev, , True, , " AL " & Format(ldFecha, "dd/mm/yyyy"))
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CLIENTE:::" & psNombreCliente & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "CREDITO:::" & psCtaCod & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CODIGO", 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("EMPR.FINAN.", 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("DEUDA IFI", 20) & Space(10)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("RELACIÓN", 40) & Space(10)
    vsCadImpresion = vsCadImpresion & lsSaltoLin & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin
    lnValorDeuda = 0

Do While Not oRs.EOF
    If oRs!bActivo = 1 Then
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Cod_Emp, 5) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!Nombre, 30) & Space(5)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!nSaldoDeuda, 10) & Space(20)
    vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(oRs!cTit_titular, 12) & Space(5)
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    lnValorDeuda = lnValorDeuda + IIf(IsNull(oRs!nSaldoDeuda), 0, oRs!nSaldoDeuda)
    End If
    oRs.MoveNext
Loop
vsCadImpresion = vsCadImpresion & String(100, "-") & lsSaltoLin

Set oFecha = Nothing
ImprimeCredDeudaSistemaFinanciero = vsCadImpresion
End Function
'FIN CTI3*********************************************************************************************
'JIPR 22/06/2018
Public Function ImprimeBoletaPagoLote(ByVal pnMovNro As String, ByVal psDescrip As String, ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psCodAge As String, ByVal nSumaTotal As Double) As String
Dim oDCredPagoDet As COMDCredito.DCOMCredito
Dim oDConsSis As COMDConstSistema.DCOMGeneral
Dim R As ADODB.Recordset
Dim F As ADODB.Recordset
Dim cInstitucion As String
Dim cTpFormaPago As String
Dim nTpFormaPago As Integer
Dim nNroDoc As String
Dim cAgencia As String
Dim dFechaHora As String
Dim i As Integer
    On Error GoTo ErrorImprimeBoleta
    
    vsCadImpresion = ""
        
    Set oDCredPagoDet = New COMDCredito.DCOMCredito
    Set oDConsSis = New COMDConstSistema.DCOMGeneral
    Set R = oDCredPagoDet.RecuperaPagoLoteDetalle(pnMovNro, psCodAge)
    
    
    Set oDCredPagoDet = Nothing
    
      If R.RecordCount > 0 Then
         If Not (R.BOF And R.EOF) Then
        cInstitucion = R!cPersNombre
        cTpFormaPago = R!cTpFormaPago
        nTpFormaPago = R!TipFormaPago
        nNroDoc = R!cNroDoc
        cAgencia = R!cAgeDescripcion
        dFechaHora = R!dFechaHora
        
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(20)  'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(76)   'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72) ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & ""

         vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat("CMAC MAYNAS" & String(3, "-") & "CREDITOS", 50)
        vsCadImpresion = vsCadImpresion & Space(5) & oImpre.ImpreFormat("CMAC MAYNAS" & String(3, "-") & "CREDITOS", 50) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(cAgencia, 35)
        vsCadImpresion = vsCadImpresion & Space(12) & oImpre.ImpreFormat(cAgencia, 50) & lsSaltoLin
                                                             
        vsCadImpresion = vsCadImpresion & Space(2) & "Fecha :" & oImpre.ImpreFormat(Format(pdFecSis, "DD/MM/YYYY"), 12) & Space(5) & "Hora : " & oImpre.ImpreFormat(Format(dFechaHora, "hh:mm:ss"), 8)
        vsCadImpresion = vsCadImpresion & Space(6) & "Fecha :" & oImpre.ImpreFormat(Format(pdFecSis, "DD/MM/YYYY"), 12) & Space(5) & "Hora : " & oImpre.ImpreFormat(Format(dFechaHora, "hh:mm:ss"), 8) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Space(2) & "Medio de Pago :" & oImpre.ImpreFormat(cTpFormaPago, 9)
        vsCadImpresion = vsCadImpresion & Space(23) & "Medio de Pago :" & oImpre.ImpreFormat(cTpFormaPago, 9) & lsSaltoLin
    
        If nTpFormaPago = 2 Then
            vsCadImpresion = vsCadImpresion & Space(2) & "Nº Cheque :" & Space(4) & oImpre.ImpreFormat(nNroDoc, 10)
            vsCadImpresion = vsCadImpresion & Space(22) & "Nº Cheque :" & Space(4) & oImpre.ImpreFormat(nNroDoc, 10) & lsSaltoLin
        End If

        vsCadImpresion = vsCadImpresion & Space(2) & "Convenio Extornado :"
        vsCadImpresion = vsCadImpresion & Space(29) & "Convenio Extornado :" & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Space(0) & oImpre.ImpreFormat(cInstitucion, 47)
        vsCadImpresion = vsCadImpresion & Space(0) & oImpre.ImpreFormat(cInstitucion, 100) & lsSaltoLin
      
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(10, "-") & psDescrip & String(10, "-"), 47)
        vsCadImpresion = vsCadImpresion & oImpre.ImpreFormat(String(10, "-") & psDescrip & String(10, "-"), 50) & lsSaltoLin

        vsCadImpresion = vsCadImpresion & Space(2) & "Monto  :" & Space(16) & oImpre.ImpreFormat(nSumaTotal, 10)
        vsCadImpresion = vsCadImpresion & Space(12) & "Monto   :" & Space(13) & oImpre.ImpreFormat(nSumaTotal, 12) & lsSaltoLin
        
        vsCadImpresion = vsCadImpresion & Space(2) & String(44, "-")
        vsCadImpresion = vsCadImpresion & Space(5) & String(42, "-") & lsSaltoLin
    
        vsCadImpresion = vsCadImpresion & Space(2) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(nSumaTotal, 12)
        vsCadImpresion = vsCadImpresion & Space(12) & "Total Pagado :" & Space(8) & oImpre.ImpreFormat(nSumaTotal, 12) & lsSaltoLin
        vsCadImpresion = vsCadImpresion & Space(2) & "Usuario Exto :" & Space(16) & oImpre.ImpreFormat(psCodUser, 12)
        vsCadImpresion = vsCadImpresion & Space(5) & "Usuario Exto :" & Space(16) & oImpre.ImpreFormat(psCodUser, 12) & lsSaltoLin


    ImprimeBoletaPagoLote = vsCadImpresion
    Exit Function
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    End If
    
    Else
         MsgBox "No se encontraron Datos", vbInformation, "Aviso"
    End If
    
        
End Function

'JIPR 22/06/2018
'CTI3 - FERIMORO - 10/12/2018 -
Private Sub ImprimeCabeceraBoletaAmpliado(ByVal sDatosBolAmpl As Variant)

Dim tipocredito As String
Dim sMax As Integer
Dim sITFdesembolso, sITFcancelado As String
'sSep = 15
'sIni = 1
sMax = 33
'saux = 5

        On Error GoTo ErrorImprimeCabeceraBoletaPagoAmpliado
        'oITF.gnITFPorcent
        'ITF por desembolso
        sITFdesembolso = calculaITF(sDatosBolAmpl(14))
        'ITF por cancelación
        sITFcancelado = calculaITF(IIf(sDatosBolAmpl(11) = "", 0, sDatosBolAmpl(11)))
      '  Set oITF = Nothing
        tipocredito = ""
        vsCadImpreAmpli = ""
        vsCadImpreAmpli = vsCadImpreAmpli & lsNegritaOn
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat(sDatosBolAmpl(8) & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(tipocredito, 1, 25), 51, 0), 51, 0) & Trim(oImpre.ImpreFormat(sDatosBolAmpl(8) & " - CREDITOS" & Space(5) & oImpre.ImpreFormat(Mid(tipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(tipocredito, 1, 25), 51, 0), 50, 0) & Space(1) & Trim(oImpre.ImpreFormat("RUC 20103845328" & Space(5) & "Cod. Ope.: 3234" & oImpre.ImpreFormat(Mid(tipocredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat(Mid(sDatosBolAmpl(1), 1, 22) & " - " & sDatosBolAmpl(2) & Space(3) & Trim(sDatosBolAmpl(6)), 51, 0) & Trim(oImpre.ImpreFormat(Mid(sDatosBolAmpl(1), 1, 22) & " - " & sDatosBolAmpl(2) & Space(3) & Trim(sDatosBolAmpl(6)), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & lsNegritaOff
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Fecha : " & sDatosBolAmpl(3) & Space(5) & "Hora : " & sDatosBolAmpl(4), 51, 0) & Trim(oImpre.ImpreFormat("Fecha : " & sDatosBolAmpl(3) & Space(5) & "Hora : " & sDatosBolAmpl(4), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Cliente : " & Mid(sDatosBolAmpl(5), 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Cliente : " & Mid(sDatosBolAmpl(5), 1, 50) & Space(1), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("DNI     : " & Mid(sDatosBolAmpl(10), 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("DNI : " & Mid(sDatosBolAmpl(10), 1, 50) & Space(1), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Crédito Vigente    : " & Mid(sDatosBolAmpl(7), 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Crédito Vigente    : " & Mid(sDatosBolAmpl(7), 1, 50) & Space(1), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Crédito a Cancelar : " & Mid(sDatosBolAmpl(12), 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Crédito a Cancelar : " & Mid(sDatosBolAmpl(12), 1, 50) & Space(1), 53, 0)) & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Cuenta : " & Mid(sDatosBolAmpl(13), 1, 50) & Space(1), 51, 0) & Trim(oImpre.ImpreFormat("Cuenta : " & Mid(sDatosBolAmpl(13), 1, 50) & Space(1), 53, 0)) & lsSaltoLin & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-")
        vsCadImpreAmpli = vsCadImpreAmpli & Space(6) & String(9, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
        '----------------------------------------------------------------------------------------------------------------------------------------------
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Monto del Desembolso" & Space(sMax + 10 - Len("Monto del Desembolso") - Len(Format(sDatosBolAmpl(14), "##,##0.00"))) & Format(sDatosBolAmpl(14), "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Monto del Desembolso" & Space(sMax + 10 - Len("Monto del Desembolso") - Len(Format(sDatosBolAmpl(14), "##,##0.00"))) & Format(sDatosBolAmpl(14), "##,##0.00"), 51, 0) & Chr(13) ' & lsSaltoLin
        Dim num As Double
        If sDatosBolAmpl(14) = "" Then sDatosBolAmpl(14) = 0
        If sDatosBolAmpl(11) = "" Then sDatosBolAmpl(11) = 0
        If sDatosBolAmpl(15) = "" Then sDatosBolAmpl(15) = 0
        'vsCadImpreAmpli = vsCadImpreAmpli & Chr(8)
        num = IIf(CDbl(sDatosBolAmpl(14)) > 0, CDbl(sDatosBolAmpl(14)), 0) - (IIf(CDbl(sDatosBolAmpl(11)) > 0, CDbl(sDatosBolAmpl(11)), 0) + IIf(CDbl(sDatosBolAmpl(15)) > 0, CDbl(sDatosBolAmpl(15)), 0))
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Depósito a CTA" & Space(sMax + 10 - Len("Depósito a CTA") - Len(Format(sDatosBolAmpl(14), "##,##0.00"))) & Format(sDatosBolAmpl(14), "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Depósito a CTA" & Space(sMax + 10 - Len("Depósito a CTA") - Len(Format(sDatosBolAmpl(14), "##,##0.00"))) & Format(sDatosBolAmpl(14), "##,##0.00"), 51, 0) & Chr(13) '& lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("ITF" & Space(sMax + 10 - Len("ITF") - Len(Format(sITFdesembolso, "##,##0.00"))) & Format(sITFdesembolso, "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("ITF" & Space(sMax + 10 - Len("ITF") - Len(Format(sITFdesembolso, "##,##0.00"))) & Format(sITFdesembolso, "##,##0.00"), 51, 0) & lsSaltoLin & Chr(13)
        
        vsCadImpreAmpli = vsCadImpreAmpli & String(16, "-") & "RETIRO AHORROS" & String(15, "-")
        vsCadImpreAmpli = vsCadImpreAmpli & Space(6) & String(16, "-") & "RETIRO AHORROS" & String(14, "-") & lsSaltoLin
        
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Retiro x Canc. Crédito" & Space(sMax + 10 - Len("Retiro x Canc. Crédito") - Len(Format(sDatosBolAmpl(11), "##,##0.00"))) & Format(sDatosBolAmpl(11), "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Retiro x Canc. Crédito" & Space(sMax + 10 - Len("Retiro x Canc. Crédito") - Len(Format(sDatosBolAmpl(11), "##,##0.00"))) & Format(sDatosBolAmpl(11), "##,##0.00"), 51, 0) & Chr(13)  '& lsSaltoLin

        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("ITF" & Space(sMax + 10 - Len("ITF") - Len(Format(sITFcancelado, "##,##0.00"))) & Format(sITFcancelado, "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("ITF" & Space(sMax + 10 - Len("ITF") - Len(Format(sITFcancelado, "##,##0.00"))) & Format(sITFcancelado, "##,##0.00"), 51, 0) & lsSaltoLin & Chr(13)

        vsCadImpreAmpli = vsCadImpreAmpli & String(45, "-")
        vsCadImpreAmpli = vsCadImpreAmpli & Space(6) & String(45, "-") & lsSaltoLin
        
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Saldo Disponible" & Space(sMax + 10 - Len("Saldo Disponible") - Len(Format(num, "##,##0.00"))) & Format(num, "##,##0.00"), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Saldo Disponible" & Space(sMax + 10 - Len("Saldo Disponible") - Len(Format(num, "##,##0.00"))) & Format(num, "##,##0.00"), 51, 0) & Chr(13) '& lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Usuario" & Space(sMax + 10 - Len("Usuario") - Len(sDatosBolAmpl(17))) & sDatosBolAmpl(17), 51, 0)
        vsCadImpreAmpli = vsCadImpreAmpli & oImpre.ImpreFormat("Usuario" & Space(sMax + 10 - Len("Usuario") - Len(sDatosBolAmpl(17))) & sDatosBolAmpl(17), 51, 0) & Chr(13)  '& lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & String(45, "-") & Space(6) & String(45, "-")
        vsCadImpreAmpli = vsCadImpreAmpli & lsSaltoLin
        vsCadImpreAmpli = vsCadImpreAmpli & Chr(12)
        
        Exit Sub

ErrorImprimeCabeceraBoletaPagoAmpliado:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub

Public Function calculaITF(ByVal nValor As Double) As Double
Dim oITF As COMDConstSistema.FCOMITF
Set oITF = New COMDConstSistema.FCOMITF

Dim nTemp, nRedondeo As Double
Dim lnPos As Integer
Dim DifITF As Double

oITF.fgITFParametros
nTemp = oITF.fgITFDesembolso(nValor)

    lnPos = InStr(1, Trim(Str(nTemp)), ".")
    If lnPos > 0 Then
        DifITF = Round((nTemp * 100 Mod 10) / 100, 2)
        DifITF = IIf(DifITF = 0.05, 0, IIf(DifITF > 0.05, DifITF - 0.05, DifITF))
    Else
        DifITF = 0
    End If

calculaITF = nTemp - DifITF

Set oITF = Nothing

End Function

