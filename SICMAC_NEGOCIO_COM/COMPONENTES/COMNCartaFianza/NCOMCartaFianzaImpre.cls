VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCartaFianzaImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function nPrintRegistroSolicitudCF(ByVal psCtaCod As String, ByVal psNomCmac As String, ByVal psNomAgencia As String, _
        ByVal psCodUser As String, ByVal psFechaSis As String) As String
Dim lsSql As String
Dim lrs As New ADODB.Recordset
Dim loCon As COMConecta.DCOMConecta

Dim lsCadImp As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas


Dim lsTipoPersonaJuridica As String, lsMagnitudEmpresarial As String, lsSiglas As String

    lsSql = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto, CE.dVenc, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & "), " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad   " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & gColocEstSolic
    Set loCon = New COMConecta.DCOMConecta
    loCon.AbreConexion
        Set lrs = loCon.CargaRecordSet(lsSql)
    loCon.CierraConexion

'lsSQL = "SELECT C.cCodCta, C.cEstado, C.dAsignacion, C.nMontosol, C.dVencSol, C.cCondCF, C.cCodAnalista, C.cModalidad, " _
'    & " C.cModalidad, " _
'    & " cNomAnalista = (Select cNomUsu From " & gcCentralCom & "Usuario WHERE cCodUsu = cCodAnalista ), " _
'    & " cTipoProd = (Case When substring(C.ccodcta,3,1) = '1' then 'COMERCIAL' " _
'    & "                   When substring(C.ccodcta,3,1) = '2' then 'MICROEMPRESARIAL' " _
'    & "                   When substring(C.ccodcta,3,1) = '3' then 'CONSUMO' Else '' End ), " _
'    & " PTi.cCodPers as cCodPersTi, PTi.cNomPers cNomPersTi, PTi.cTipPers, PTi.cTidoci cTidociTi, PTi.cNudoci cNudociTi, " _
'    & " PTi.cTidotr as cTidoTriTi , PTi.cNudotr cNudoTriTi, " _
'    & " PTi.cDirPers cDirPersTi, PTi.cCodZon as ZonaDirTi, Z1.cDesZon as cDesZonDirTi, PTi.dFecNac dFecNacTi, " _
'    & " cTipoPersTi = (Case When PTi.cTipPers ='1' then 'PERSONA NATURAL' " _
'    & "                     When PTi.cTipPers ='2' then 'PERSONA JURIDICA'  Else '' End ) , " _
'    & " PAc.cCodPers as cCodPersAc, PAc.cNomPers as cNomPersAc, PAc.cTipPers as cTipPersAc, PAc.cTidoci cTidociAc, PAc.cNudoci cNudociAc, PAc.cTidotr cTidotriAc, PAc.cNudotr cNudoTriAc, " _
'    & " PAc.cDirPers cDirPersAc, PAc.cCodZon as ZonaDirAc, Z2.cDesZon as cDesZonDirAc, PAc.dFecNac, " _
'    & " FI.cTipoFuente, FI.cRazonSocial, FI.cTidotri, " _
'    & " FI.cCodZon as ZonaFuente, FI.cDireccion, FI.cTelefono, FI.cSector, " _
'    & " FI.cActEcon, FI.nNroEmp, FI.cTiemAct, FI.cCargo " _
'    & " FROM CartaFianza C " _
'    & " INNER JOIN PersCartaFianza PCTi ON C.cCodCta = PCTi.cCodCta " _
'    & " INNER JOIN " & gcCentralPers & "Persona PTi ON PCTi.cCodPers = PTi.cCodPers " _
'    & " INNER JOIN PersCartaFianza PCAc ON C.cCodCta = PCAc.cCodCta " _
'    & " INNER JOIN " & gcCentralPers & "Persona PAc ON PCAc.cCodPers = PAc.cCodPers " _
'    & " INNER JOIN FuenteIngreso FI ON C.cNumFuente = FI.cNumFuente " _
'    & " INNER JOIN " & gcCentralCom & "Zonas Z1 ON  Z1.cCodZon = PTi.cCodZon " _
'    & " INNER JOIN " & gcCentralCom & "Zonas Z2 ON  Z2.cCodZon = PAc.cCodZon  " _
'    & " WHERE C.cCodCta='" & psCodCta & "' AND PCTi.cRelaCta ='TI' AND PCAc.cRelaCta ='AC'  "
'
'Set lrs = CargaRecord(lsSQL)
'
'lsImp = ""
'If RSVacio(lrs) Then
'    MsgBox "No encuentra información sobre la Carta Fianza", vbInformation, "Aviso"
'    Exit Function
'Else
'
'
'    lsImp = lsImp & gsNomCmac & Space(65 - Len(gsNomCmac)) & "Fecha : " & Format(gdFecSis, "dd/mm/yyyy") & gPrnSaltoLinea
'    lsImp = lsImp & gsNomAge & Space(65 - Len(gsNomAge)) & "Hora  : " & Format(Time(), "hh:mm:ss") & gPrnSaltoLinea
'
'    lsImp = lsImp & gPrnBoldON
'    lsImp = lsImp & Space(20) & "REGISTRO DE SOLICITUD DE CARTA FIANZA" & gPrnSaltoLinea
'    lsImp = lsImp & gPrnSaltoLinea
'    lsImp = lsImp & Space(5) & "CARTA FIANZA : " & psCodCta & Space(10) & "CARTA FIANZA " & lrs!cTipoProd & gPrnSaltoLinea
'    lsImp = lsImp & gPrnSaltoLinea
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "DATOS DEL TITULAR" & Space(25) & lrs!cTipoPersTi & gPrnSaltoLinea
'    lsImp = lsImp & gPrnBoldOFF
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "Cliente         : " & lrs!cCodPersTi & Space(2) & ImpreCarEsp(PstaNombre(lrs!cNomPersTi, False)) & gPrnSaltoLinea
'    If IsNull(lrs!cTidoCiTi) Then
'        lsImp = lsImp & "Doc. Identidad  : " & Space(10)
'    Else
'        lsImp = lsImp & "Doc. Identidad  : " & Tablacod("04", lrs!cTidoCiTi) & " " & lrs!cNudociTi & gPrnSaltoLinea
'    End If
'    If IsNull(lrs!cTidoTriTi) Then
'        lsImp = lsImp & "Doc. Tributario : " & Space(10)
'    Else
'        lsImp = lsImp & "Doc. Tributario : " & Tablacod("05", lrs!cTidoTriTi) & " " & lrs!cNudoTriTi & gPrnSaltoLinea
'    End If
'    lsImp = lsImp & "Direccion       : " & lrs!cDirPersTi & " - " & lrs!cDesZonDirTi & gPrnSaltoLinea
'
'    If lrs!cTipPers = "1" Then
'        lsSQL = "Select * From " & gcCentralPers & "PersonaNat where cCodPers='" & Trim(lrs!cCodPersTi) & "'"
'        Set lrPers = CargaRecord(lsSQL)
'        If Not RSVacio(lrPers) Then
'            lsProfesion = Tablacod("34", IIf(IsNull(lrPers!cCodProf), "", lrPers!cCodProf))
'            lsEstadoCivil = Tablacod("06", IIf(IsNull(lrPers!cEstCiv), "", lrPers!cEstCiv))
'            lsHijos = IIf(IsNull(lrPers!nNumhij), "", Format(lrPers!nNumhij, "#0"))
'            If Len(Trim(lsNombreConyugue)) = 0 Then
'                'lsNombreConyugue = DatosConyugue(Trim(IIf(IsNull(lrPers!cCodCon), "", lrPers!cCodCon)))
'            End If
'        End If
'        lrPers.Close
'        Set lrPers = Nothing
'
'        lsImp = lsImp & "Profesion      : " & lrs!cActEcon & gPrnSaltoLinea
'        'lsImp = lsImp & "Edad           : " & DateDiff("year", lrs!dFecNac, gdFecSis) & gPrnSaltoLinea
'    Else
'        '***
'        lsSQL = "Select * from " & gcCentralPers & "PersonaJur where cCodPers='" & Trim(lrs!cCodPersTi) & "'"
'        Set lrPers = CargaRecord(lsSQL)
'        If Not RSVacio(lrPers) Then
'            lsTipoPersonaJuridica = Tablacod("33", IIf(IsNull(lrPers!cTipPerJur), "", lrPers!cTipPerJur))
'            lsMagnitudEmpresarial = Tablacod("19", IIf(IsNull(lrPers!cMagEmp), "", lrPers!cMagEmp))
'            lsSiglas = IIf(IsNull(lrPers!cSigla), "", lrPers!cSigla)
'        End If
'        lrPers.Close
'        Set lrPers = Nothing
'        ' ***
'        lsImp = lsImp & "Tiempo Creacion : " & IIf(IsNull(lrs!cTiemAct), "", lrs!cTiemAct) & gPrnSaltoLinea
'        lsImp = lsImp & "Tipo Juridico   : " & lsTipoPersonaJuridica & gPrnSaltoLinea
'        lsImp = lsImp & "Magnitud Empr.  : " & lsMagnitudEmpresarial & gPrnSaltoLinea
'    End If
'
'    ' ****** Acreedor
'    lsImp = lsImp & gPrnSaltoLinea & gPrnBoldON
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "DATOS DEL ACREEDOR" & Space(25) & lrs!cTipPersAc & gPrnSaltoLinea
'    lsImp = lsImp & gPrnBoldOFF
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "Cliente         : " & lrs!cCodPersAc & Space(2) & ImpreCarEsp(PstaNombre(lrs!cNomPersAC, False)) & gPrnSaltoLinea
'    If IsNull(lrs!cTidoCiAc) Then
'        lsImp = lsImp & "Doc. Identidad  : " & Space(10)
'    Else
'        lsImp = lsImp & "Doc. Identidad  : " & Tablacod("04", lrs!cTidoCiAc) & " " & lrs!cNudociTi & gPrnSaltoLinea
'    End If
'    If IsNull(lrs!cTidoTriAc) Then
'        lsImp = lsImp & "Doc. Tributario : " & Space(10)
'    Else
'        lsImp = lsImp & "Doc. Tributario : " & Tablacod("05", lrs!cTidoTriAc) & " " & lrs!cNudoTriAc & gPrnSaltoLinea
'    End If
'    lsImp = lsImp & "Direccion       : " & lrs!cDirPersAc & " - " & lrs!cDesZonDirAc & gPrnSaltoLinea
'
'    lsImp = lsImp & gPrnSaltoLinea & gPrnBoldON
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "FUENTE DE INGRESO" & Space(20) & Tablacod("54", IIf(IsNull(lrs!cTipoFuente), "", lrs!cTipoFuente)) & gPrnBoldOFF & gPrnSaltoLinea
'    lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'    lsImp = lsImp & "Razon Social   : " & lrs!cRazonSocial & gPrnSaltoLinea
'    lsImp = lsImp & "Direccion      : " & lrs!cdireccion & gPrnSaltoLinea
'    lsImp = lsImp & "Sector         : " & Tablacod("20", IIf(IsNull(lrs!csector), "", lrs!csector)) & gPrnSaltoLinea
'    lsImp = lsImp & "Actividad      : " & Tablacod("35", IIf(IsNull(lrs!cActEcon), "", lrs!cActEcon)) & gPrnSaltoLinea
'    lsImp = lsImp & "Inicio de Actividades  : " & IIf(IsNull(lrs!cTiemAct), "", lrs!cTiemAct) & gPrnSaltoLinea
'
'    If lnTotalRelaciones > 0 Then
'        lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'        lsImp = lsImp & gPrnBoldON & Space(10) & "CLIENTES RELACIONADOS" & gPrnSaltoLinea & gPrnBoldOFF
'        lsImp = lsImp & ImpreCarEsp("  CODIGO  " & Space(2) & "NOMBRE" & Space(46) & "RELACION" & Space(4) & "DOC.IDENT.") & gPrnSaltoLinea
'        lsImp = lsImp & String(90, "-") & gPrnSaltoLinea
'        For i = 1 To lnTotalRelaciones
'            lsImp = lsImp & ImpreCarEsp(ImpreCarEsp(Relaciones(1, i)) & Space(2) & CadDerecha(Relaciones(2, i), 50) & Space(2) & ofun.ImpreFormat(Relaciones(3, i), 8) & Space(4) & Relaciones(4, i)) & gPrnSaltoLinea
'        Next i
'    End If
'
'    lsImp = lsImp & gPrnBoldON & String(90, "_") & gPrnSaltoLinea
'    lsImp = lsImp & ImpreCarEsp(CentrarCadena("DATOS DE LA CARTA FIANZA", 80) + gPrnBoldOFF) & gPrnSaltoLinea
'    lsImp = lsImp & ImpreCarEsp("Analista         : " & lrs!cNomAnalista) & gPrnSaltoLinea
'    lsImp = lsImp & ImpreCarEsp("Modalidad        : " & Tablacod("D1", IIf(IsNull(lrs!cModalidad), "", lrs!cModalidad))) & Space(15)
'    lsImp = lsImp & ImpreCarEsp("Condición        : " & Tablacod("38", IIf(IsNull(lrs!cCondCF), "", lrs!cCondCF))) & gPrnSaltoLinea
'    lsImp = lsImp & ImpreCarEsp("Fecha Asignación : " & Format(lrs!dAsignacion, "dd/mm/yyyy")) & gPrnSaltoLinea
'    lsImp = lsImp & gPrnBoldON
'    lsImp = lsImp & ImpreCarEsp("Monto Solicitado : " & ofun.ImpreFormat(lrs!nMontoSol, 10, 2, True) & " " & IIf(Mid(lrs!cCodCta, 6, 1) = "1", "NUEVOS SOLES ", "DOLARES")) & gPrnSaltoLinea
'    lsImp = lsImp & ImpreCarEsp("Vencimiento      : " & Format(lrs!dVencSol, "mm/dd/yyyy"))
'    lsImp = lsImp & gPrnSaltoLinea
'    lsImp = lsImp & String(90, "_") & gPrnBoldOFF
'    lsImp = lsImp & gPrnSaltoPagina  'expulsa la pagina
'    ImpreBegin False, 66
'    Print #ArcSal, lsImp
'    ImpreEnd
'End If
'lrs.Close
'Set lrs = Nothing
End Function

Public Function nPrintReciboCFComision(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, ByVal psNombAcreedor As String, _
        ByVal pnImporteCF As Currency, ByVal pdFecVencCF As Date, _
        ByVal pnComision As Currency, ByVal psCodUser As String, _
        ByVal psmensaje As String, Optional ByVal psCodCmac As String, _
        Optional ByVal pnITF As Double = 0#, Optional ByVal psImpresora As Impresoras = gEPSON, _
        Optional ByVal pbImpTMU As Boolean = False, Optional ByVal pnTipo As Integer = 0) As String
    'WIOR 20120806 AGREGO pnTipo PARA IDENTIFICAR EL TIPO DE COMISION
    Dim lsCadImp As String
    
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim ofcar As New COMFunciones.FCOMCadenas
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
    lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita


    Dim vEspacio As Integer
    vEspacio = 5
    
    If psCodCmac = "102" Then  '' Lima
            lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Cuenta    :  " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(psNombAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(pnImporteCF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(pdFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(20) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
            
    Else
        If pbImpTMU = False Then
            lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(11) & Space(10 + vEspacio) _
                                & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            'WIOR 20120806**************
            If pnTipo = 0 Then
            lsCadImp = lsCadImp & " Comisión por Carta Fianza " & Space(18 + vEspacio) _
                                & "Comisión por Carta Fianza " & oFunI.gPrnSaltoLinea
            ElseIf pnTipo = 1 Then
            lsCadImp = lsCadImp & " Comisión por Modificación de Carta Fianza" & Space(3 + vEspacio) _
                                & "Comisión por Modificación de Carta Fianza" & oFunI.gPrnSaltoLinea
            End If
            'WIOR FIN ******************
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(5 + vEspacio) _
                                & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Cuenta    :  " & Mid(psCodCta, 1, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10) & Space(9) & Space(vEspacio) _
                                & " Cuenta    :  " & Mid(psCodCta, 1, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10) & Space(11) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & Space(1 + vEspacio) _
                                & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(psNombAcreedor, 28) & Space(1 + vEspacio) _
                                & " ACREEDOR  : " & ofun.ImpreFormat(psNombAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(pnImporteCF, 12, 2) & Space(16) & Space(vEspacio) _
                                & " Importe   : " & ofun.ImpreFormat(pnImporteCF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(pdFecVencCF, "dd/mm/yyyy") & Space(21) & Space(vEspacio) _
                                & " Vencim.   : " & Format(pdFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & Space(4 + vEspacio) _
                                & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " I.T.F.         : " & ofun.ImpreFormat(pnITF, 12, 2) & Space(11 + vEspacio) _
                                & " I.T.F.         : " & ofun.ImpreFormat(pnITF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & Space(11 + vEspacio) _
                                & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Total a Pagar. : " & ofun.ImpreFormat(pnComision + pnITF, 12, 2) & Space(11 + vEspacio) _
                                & " Total a Pagar. : " & ofun.ImpreFormat(pnComision + pnITF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & Space(4 + vEspacio) _
                                & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(40) & Space(vEspacio) _
                                & " " & Format(psCodUser, "@@@@") & Space(20) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 1) _
                                & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = Replace(lsCadImp, "k", "") 'WIOR 20120806
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            'WIOR 20120806**************
            If pnTipo = 0 Then
            lsCadImp = lsCadImp & "Comisión - Carta Fianza" & oFunI.gPrnSaltoLinea
            ElseIf pnTipo = 1 Then
            lsCadImp = lsCadImp & "Comisión - Modificación de Carta Fianza" & oFunI.gPrnSaltoLinea
            End If
            'WIOR FIN ******************
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            'lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " CUENTA    :  " & Mid(psCodCta, 1, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10) & oFunI.gPrnSaltoLinea
            'lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(psNombAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(pnImporteCF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(pdFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            'lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " I.T.F.         : " & ofun.ImpreFormat(pnITF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Total a Pagar. : " & ofun.ImpreFormat(pnComision + pnITF, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
            'lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        End If
    End If
    nPrintReciboCFComision = lsCadImp
    Set ofun = Nothing
    Set oFunI = Nothing
    Set ofcar = Nothing
End Function


Public Function nPrintReciboCFComisionExtorno(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal pnComision As Currency, ByVal psCodUser As String, _
        ByVal psmensaje As String, Optional ByVal psCodCmac As String, _
        Optional ByVal psImpresoras As Impresoras = gEPSON, _
        Optional ByVal pbImpTMU As Boolean = False) As String
    
    Dim lsCadImp As String
    
    Dim lsSql As String
    Dim lr As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim ofcar As New COMFunciones.FCOMCadenas
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresoras
    
    
    Dim lsNomCliente As String, lsNomAcreedor As String
    Dim lnImporteCF As Double, ldFecVencCF As Date
    
    lsSql = "Select P.cPersNombre, P2.cPersNombre as cPersNomAcre, " _
        & " Prd.nSaldo as nMontoApr, CCF.dVencimiento as dVencApr " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " WHERE CCF.cCtaCod = '" & psCodCta & "' "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set lr = oCon.CargaRecordSet(lsSql)
    lsNomCliente = lr!cPersNombre:    lsNomAcreedor = lr!cPersNomAcre
    lnImporteCF = lr!nMontoApr:       ldFecVencCF = Format(lr!dVencApr, "dd/mm/yyyy")
    lr.Close
    oCon.CierraConexion
    Set oCon = Nothing

'    lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
'    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
'    lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita


    Dim vEspacio As Integer
        
    If psCodCmac = "102" Then
           vEspacio = 5
            lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Cuenta    :  " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(lsNomCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(lsNomAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(lnImporteCF, 10, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(ldFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===========EXTORNO DE COMISION======= " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(20) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
    Else
        vEspacio = 5
        If pbImpTMU = False Then
            lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & Space(11) & Space(10 + vEspacio) _
                                & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(5 + vEspacio) _
                                & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Cuenta    :  " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & Space(vEspacio) _
                                & " Cuenta    :  " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(11) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(lsNomCliente, 28) & Space(1 + vEspacio) _
                                & " CLIENTE   : " & ofun.ImpreFormat(lsNomCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(lsNomAcreedor, 28) & Space(1 + vEspacio) _
                                & " ACREEDOR  : " & ofun.ImpreFormat(lsNomAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(lnImporteCF, 10, 2) & Space(18) & Space(vEspacio) _
                                & " Importe   : " & ofun.ImpreFormat(lnImporteCF, 10, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(ldFecVencCF, "dd/mm/yyyy") & Space(21) & Space(vEspacio) _
                                & " Vencim.   : " & Format(ldFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===========EXTORNO DE COMISION======= " & Space(4 + vEspacio) _
                                & " ===========EXTORNO DE COMISION======= " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & Space(11 + vEspacio) _
                                & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & Space(4 + vEspacio) _
                                & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(40) & Space(vEspacio) _
                                & " " & Format(psCodUser, "@@@@") & Space(20) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & Space(vEspacio + 1) _
                                & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
            lsCadImp = lsCadImp & " Carta Fianza - " & IIf(Mid(psCodCta, 9, 1) = "1", "Soles  ", "Dolares") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
            lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 20, 0) & Space(2) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Cuenta    :  " & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(lsNomCliente, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ACREEDOR  : " & ofun.ImpreFormat(lsNomAcreedor, 28) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Importe   : " & ofun.ImpreFormat(lnImporteCF, 10, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " Vencim.   : " & Format(ldFecVencCF, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===========EXTORNO DE COMISION======= " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
            lsCadImp = lsCadImp & " Comision  C.F. : " & ofun.ImpreFormat(pnComision, 12, 2) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " ===================================== " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(40) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
            lsCadImp = lsCadImp & "" & ofun.ImpreFormat(psmensaje, 40) & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        End If
    End If
        nPrintReciboCFComisionExtorno = lsCadImp
        Set oFunI = Nothing
End Function




